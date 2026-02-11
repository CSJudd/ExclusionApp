import sqlite3
import pandas as pd
from pathlib import Path
from datetime import datetime

from engine.config_loader import ClientConfig
from engine.normalizer import (
    normalize_person_name,
    normalize_entity_name,
    normalize_dob,
    normalize_zip,
    extract_ssn_last4
)
from engine.vendor_classifier import classify_vendor
from engine.matcher_people import match_person
from engine.matcher_entity import match_entity
from engine.reference_cache import get_cache_path, file_sha256
from engine.audit_xlsx import write_audit_workbook
from engine.history import create_run_directory, write_metadata, write_run_log
from engine.version import ENGINE_VERSION, MATCH_THRESHOLD_VERSION
from engine.pdf_reports import generate_pdf_report

def run_exclusion_check(
    client_config_path,
    month,
    staff_path=None,
    board_path=None,
    vendor_path=None,
    oig_path=None,
    sam_path=None
):
    config = ClientConfig(client_config_path)
    db_path = get_cache_path(month)

    if not db_path.exists():
        raise Exception(f"Reference cache for {month} does not exist.")

    conn = sqlite3.connect(db_path)

    results = {
        "staff": [],
        "board": [],
        "vendors": []
    }

    # ---------------- STAFF ----------------
    if staff_path:
        staff_section = config.section("staff")
        df_staff = pd.read_csv(staff_path, dtype=str).fillna("")

        for _, row in df_staff.iterrows():
            first, last, middle, full = normalize_person_name(
                row.get(staff_section["first_name"]),
                row.get(staff_section["last_name"]),
                row.get(staff_section.get("middle_name"))
            )

            dob_iso, dob_compact = normalize_dob(row.get(staff_section["dob"]))
            zip_code = normalize_zip(row.get(staff_section.get("zip")))
            ssn_last4 = extract_ssn_last4(row.get(staff_section.get("ssn")))

            match_result = match_person(
                conn,
                first,
                last,
                dob_compact=dob_compact,
                zip_code=zip_code
            )

            results["staff"].append({
                "name": full,
                "dob": dob_iso,
                "ssn_last4": ssn_last4,
                "role": row.get(staff_section.get("job_title")),
                "status": row.get(staff_section.get("status")),
                **match_result
            })

    # ---------------- BOARD ----------------
    if board_path:
        board_section = config.section("board")
        skip_rows = board_section.get("skip_rows", 0)
        df_board = pd.read_csv(board_path, dtype=str, skiprows=skip_rows).fillna("")

        for _, row in df_board.iterrows():
            full_name = row.get(board_section["name_column"], "")
            tokens = full_name.split()
            first = tokens[0] if tokens else ""
            last = tokens[-1] if len(tokens) > 1 else ""

            first, last, middle, full = normalize_person_name(first, last)

            dob_iso, dob_compact = normalize_dob(row.get(board_section["dob"]))
            zip_code = normalize_zip(row.get(board_section.get("zip")))
            ssn_last4 = extract_ssn_last4(row.get(board_section.get("ssn")))

            match_result = match_person(
                conn,
                first,
                last,
                dob_compact=dob_compact,
                zip_code=zip_code
            )

            results["board"].append({
                "name": full,
                "dob": dob_iso,
                "ssn_last4": ssn_last4,
                **match_result
            })

    # ---------------- VENDORS ----------------
    if vendor_path:
        vendor_section = config.section("vendors")
        df_vendor = pd.read_excel(vendor_path, dtype=str).fillna("")

        for _, row in df_vendor.iterrows():
            name_raw = row.get(vendor_section["entity_name"])
            tax_id = row.get(vendor_section.get("tax_id"))

            classification = classify_vendor(name_raw, tax_id)

            normalized_entity = normalize_entity_name(name_raw)
            zip_code = normalize_zip(row.get(vendor_section.get("zip")))
            state = row.get(vendor_section.get("state"), "").upper()

            if classification == "ENTITY":
                match_result = match_entity(
                    conn,
                    normalized_entity,
                    state=state,
                    zip_code=zip_code
                )
            elif classification == "PERSON_VENDOR":
                tokens = name_raw.split()
                first = tokens[0] if tokens else ""
                last = tokens[-1] if len(tokens) > 1 else ""
                first, last, middle, full = normalize_person_name(first, last)

                match_result = match_person(
                    conn,
                    first,
                    last,
                    zip_code=zip_code
                )
            else:
                entity_result = match_entity(
                    conn,
                    normalized_entity,
                    state=state,
                    zip_code=zip_code
                )

                tokens = name_raw.split()
                first = tokens[0] if tokens else ""
                last = tokens[-1] if len(tokens) > 1 else ""
                first, last, middle, full = normalize_person_name(first, last)

                person_result = match_person(
                    conn,
                    first,
                    last,
                    zip_code=zip_code
                )

                match_result = entity_result
                if (
                    person_result["oig_status"] == "CONFIRMED"
                    or person_result["sam_status"] == "CONFIRMED"
                ):
                    match_result = person_result

            results["vendors"].append({
                "name": name_raw,
                "classification": classification,
                **match_result
            })

    conn.close()

    # ---------------- HISTORY + AUDIT ----------------
    run_dir = create_run_directory(config.client_name, month)

    audit_path = run_dir / "Audit.xlsx"

    metadata = {
        "client": config.client_name,
        "month": month,
        "engine_version": ENGINE_VERSION,
        "threshold_version": MATCH_THRESHOLD_VERSION,
        "timestamp": datetime.utcnow().isoformat(),
        "staff_count": len(results["staff"]),
        "board_count": len(results["board"]),
        "vendor_count": len(results["vendors"]),
    }

    if oig_path:
        metadata["oig_file_hash"] = file_sha256(oig_path)
    if sam_path:
        metadata["sam_file_hash"] = file_sha256(sam_path)

    write_audit_workbook(
        audit_path,
        config.client_name,
        month,
        results,
        metadata
    )

    write_metadata(run_dir, metadata)
    write_run_log(run_dir, "Exclusion check completed successfully.")

    # ---------------- PDF REPORTS ----------------
    staff_pdf = run_dir / "Staff_Report.pdf"
    board_pdf = run_dir / "Board_Report.pdf"
    vendor_pdf = run_dir / "Vendor_Report.pdf"

    generate_pdf_report(
        staff_pdf,
        config.client_name,
        month,
        "Staff Exclusion Report",
        results["staff"]
    )

    generate_pdf_report(
        board_pdf,
        config.client_name,
        month,
        "Board Exclusion Report",
        results["board"]
    )

    generate_pdf_report(
        vendor_pdf,
        config.client_name,
        month,
        "Vendor Exclusion Report",
        results["vendors"]
    )

    return {
        "run_directory": str(run_dir),
        "audit_file": str(audit_path),
        "staff_pdf": str(staff_pdf),
        "board_pdf": str(board_pdf),
        "vendor_pdf": str(vendor_pdf)
    }
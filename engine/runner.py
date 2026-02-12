import sqlite3
import shutil
from pathlib import Path
from datetime import datetime

import pandas as pd

from engine.config_loader import ClientConfig
from engine.normalizer import (
    normalize_person_name,
    normalize_entity_name,
    normalize_dob,
    normalize_zip,
    extract_ssn_last4,
)
from engine.vendor_classifier import classify_vendor
from engine.matcher_people import match_person
from engine.matcher_entity import match_entity
from engine.reference_cache import get_cache_path, file_sha256
from engine.audit_xlsx import write_audit_workbook
from engine.history import create_run_directory, write_metadata, write_run_log
from engine.pdf_reports import generate_pdf_report

# NOTE: project currently provides these at repo root (not engine.version)
from version import ENGINE_VERSION, MATCH_THRESHOLD_VERSION


def _strip_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _row_get(row, key: str):
    """
    Robust column getter:
    - exact match
    - stripped match
    - case-insensitive match
    Returns "" if missing (and never returns None).
    """
    if not key:
        return ""
    key_stripped = str(key).strip()
    if key_stripped in row:
        val = row.get(key_stripped, "")
        return "" if val is None else val

    lower_map = {str(k).strip().casefold(): k for k in row.keys()}
    k2 = lower_map.get(key_stripped.casefold())
    if k2 is not None:
        val = row.get(k2, "")
        return "" if val is None else val

    return ""


def _resolve_column(df: pd.DataFrame, requested: str | None, fallbacks: list[str]) -> str:
    """
    Resolve a column name against a dataframe even when headers have whitespace/case differences.
    If requested is missing/blank or not found, try fallbacks by exact/contains matching.
    Returns the actual dataframe column name.
    Raises ValueError if no suitable column can be found.
    """
    cols = [str(c) for c in df.columns]
    cols_cf = {str(c).strip().casefold(): str(c) for c in cols}

    def has(colname: str) -> str | None:
        if not colname:
            return None
        key = str(colname).strip().casefold()
        return cols_cf.get(key)

    # 1) requested exact
    if requested:
        resolved = has(requested)
        if resolved:
            return resolved

    # 2) fallbacks exact
    for fb in fallbacks:
        resolved = has(fb)
        if resolved:
            return resolved

    # 3) fallbacks "contains" (casefold)
    cols_cf_list = [str(c).strip().casefold() for c in cols]
    for fb in fallbacks:
        fb_cf = str(fb).strip().casefold()
        for i, ccf in enumerate(cols_cf_list):
            if fb_cf in ccf:
                return cols[i]

    raise ValueError(f"Could not resolve column. Requested={requested!r}. Columns={cols!r}")


def _detect_vendor_header_row(vendor_path) -> int | None:
    """
    Some client vendor workbooks contain title/date rows above the real header.
    Detect the first row that appears to be the NAME/CITY/STATE header row.
    """
    try:
        preview = pd.read_excel(vendor_path, dtype=str, header=None, nrows=40).fillna("")
    except Exception:
        return None

    def norm(v: str) -> str:
        return str(v or "").strip().casefold()

    name_tokens = {"name", "vendor", "vendor name", "payee", "entity"}
    city_tokens = {"city", "town", "municipality"}
    state_tokens = {"state", "st", "province"}

    for idx, row in preview.iterrows():
        values = {norm(v) for v in row.tolist() if str(v).strip()}
        if not values:
            continue

        has_name = any(v in name_tokens for v in values)
        has_city = any(v in city_tokens for v in values)
        has_state = any(v in state_tokens for v in values)
        if has_name and has_city and has_state:
            return int(idx)

    return None


def run_exclusion_check(
    client_config_path,
    month,
    staff_path=None,
    board_path=None,
    vendor_path=None,
    oig_path=None,
    sam_path=None,
):
    config = ClientConfig(client_config_path)
    db_path = get_cache_path(month)

    if not db_path.exists():
        raise Exception(f"Reference cache for {month} does not exist.")

    conn = sqlite3.connect(db_path)

    results = {"staff": [], "board": [], "vendors": []}
    review_counter = 1

    # ---------------- STAFF ----------------
    if staff_path:
        staff_section = config.section("staff")
        df_staff = pd.read_csv(staff_path, dtype=str).fillna("")
        df_staff = _strip_columns(df_staff)

        for _, row in df_staff.iterrows():
            first, last, middle, full = normalize_person_name(
                _row_get(row, staff_section["first_name"]),
                _row_get(row, staff_section["last_name"]),
                _row_get(row, staff_section.get("middle_name")),
            )

            dob_iso, dob_compact = normalize_dob(_row_get(row, staff_section["dob"]))
            zip_code = normalize_zip(_row_get(row, staff_section.get("zip")))
            ssn_last4 = extract_ssn_last4(_row_get(row, staff_section.get("ssn")))

            match_result = match_person(
                conn,
                first,
                last,
                dob_compact=dob_compact,
                zip_code=zip_code,
            )

            results["staff"].append(
                {
                    "review_id": f"R{review_counter:05d}",
                    "category": "staff",
                    "name": full,
                    "last_name_display": str(_row_get(row, staff_section["last_name"])).strip(),
                    "first_name_display": " ".join(
                        p
                        for p in [
                            str(_row_get(row, staff_section["first_name"])).strip(),
                            str(_row_get(row, staff_section.get("middle_name"))).strip(),
                        ]
                        if p
                    ),
                    "name_display": " ".join(
                        p
                        for p in [
                            str(_row_get(row, staff_section["first_name"])).strip(),
                            str(_row_get(row, staff_section.get("middle_name"))).strip(),
                            str(_row_get(row, staff_section["last_name"])).strip(),
                        ]
                        if p
                    ),
                    "dob": dob_iso,
                    "ssn_last4": ssn_last4,
                    "role": _row_get(row, staff_section.get("job_title")),
                    "status": _row_get(row, staff_section.get("status")),
                    **match_result,
                }
            )
            review_counter += 1

    # ---------------- BOARD ----------------
    if board_path:
        board_section = config.section("board")
        skip_rows = board_section.get("skip_rows", 0)
        df_board = pd.read_csv(board_path, dtype=str, skiprows=skip_rows).fillna("")
        df_board = _strip_columns(df_board)

        for _, row in df_board.iterrows():
            full_name = _row_get(row, board_section["name_column"])
            tokens = str(full_name).split()
            first = tokens[0] if tokens else ""
            last = tokens[-1] if len(tokens) > 1 else ""
            first, last, middle, full = normalize_person_name(first, last)

            dob_iso, dob_compact = normalize_dob(_row_get(row, board_section["dob"]))
            zip_code = normalize_zip(_row_get(row, board_section.get("zip")))
            ssn_last4 = extract_ssn_last4(_row_get(row, board_section.get("ssn")))

            match_result = match_person(
                conn,
                first,
                last,
                dob_compact=dob_compact,
                zip_code=zip_code,
            )

            results["board"].append(
                {
                    "review_id": f"R{review_counter:05d}",
                    "category": "board",
                    "name": full,
                    "name_display": str(full_name).strip(),
                    "dob": dob_iso,
                    "ssn_last4": ssn_last4,
                    "role": "",
                    "status": "",
                    **match_result,
                }
            )
            review_counter += 1

    # ---------------- VENDORS ----------------
    df_vendor = None
    vendor_header_row = None
    entity_col = ""
    city_col = None
    state_col = None
    if vendor_path:
        vendor_section = config.section("vendors")
        vendor_header_row = _detect_vendor_header_row(vendor_path)
        if vendor_header_row is not None:
            df_vendor = pd.read_excel(vendor_path, dtype=str, header=vendor_header_row).fillna("")
        else:
            df_vendor = pd.read_excel(vendor_path, dtype=str).fillna("")
        df_vendor = _strip_columns(df_vendor)

        # Resolve configured columns against real Excel headers (whitespace/case drift safe)
        try:
            entity_col = _resolve_column(
                df_vendor,
                vendor_section.get("entity_name"),
                fallbacks=[
                    "Vendor",
                    "Vendor Name",
                    "VendorName",
                    "Name",
                    "Payee",
                    "Payee Name",
                    "Supplier",
                    "Supplier Name",
                    "Company",
                    "Company Name",
                    "Entity",
                    "Entity Name",
                ],
            )
        except Exception as e:
            raise Exception(f"Vendor entity_name column mapping failed: {e}")

        taxid_col = None
        if vendor_section.get("tax_id"):
            try:
                taxid_col = _resolve_column(
                    df_vendor,
                    vendor_section.get("tax_id"),
                    fallbacks=["TIN", "Tax ID", "TaxID", "EIN", "SSN", "TIN/EIN", "FEIN"],
                )
            except Exception:
                taxid_col = None

        try:
            state_col = _resolve_column(
                df_vendor,
                vendor_section.get("state"),
                fallbacks=["State", "ST", "STATE", "Province"],
            )
        except Exception:
            state_col = None

        try:
            city_col = _resolve_column(
                df_vendor,
                vendor_section.get("city"),
                fallbacks=["City", "CITY", "Town", "Municipality"],
            )
        except Exception:
            city_col = None

        zip_col = None
        if vendor_section.get("zip"):
            try:
                zip_col = _resolve_column(
                    df_vendor,
                    vendor_section.get("zip"),
                    fallbacks=["Zip", "ZIP", "Zip Code", "ZIP Code", "Postal", "Postal Code"],
                )
            except Exception:
                zip_col = None

        for _, row in df_vendor.iterrows():
            name_raw = _row_get(row, entity_col)
            if not str(name_raw).strip():
                continue
            name_raw_cf = str(name_raw).strip().casefold()
            if name_raw_cf in {"name", "vendor", "vendor name", "entity", "payee"}:
                continue

            tax_id = _row_get(row, taxid_col) if taxid_col else ""
            classification = classify_vendor(name_raw, tax_id)

            normalized_entity = normalize_entity_name(name_raw)
            zip_code = normalize_zip(_row_get(row, zip_col) if zip_col else "")
            state = str(_row_get(row, state_col) if state_col else "").upper()

            if classification == "ENTITY":
                match_result = match_entity(conn, normalized_entity, state=state, zip_code=zip_code)

            elif classification == "PERSON_VENDOR":
                tokens = str(name_raw).split()
                first = tokens[0] if tokens else ""
                last = tokens[-1] if len(tokens) > 1 else ""
                first, last, middle, full = normalize_person_name(first, last)
                match_result = match_person(conn, first, last, zip_code=zip_code)

            else:
                entity_result = match_entity(conn, normalized_entity, state=state, zip_code=zip_code)

                tokens = str(name_raw).split()
                first = tokens[0] if tokens else ""
                last = tokens[-1] if len(tokens) > 1 else ""
                first, last, middle, full = normalize_person_name(first, last)
                person_result = match_person(conn, first, last, zip_code=zip_code)

                match_result = entity_result
                if person_result.get("oig_status") == "CONFIRMED" or person_result.get("sam_status") == "CONFIRMED":
                    match_result = person_result

            results["vendors"].append(
                {
                    "review_id": f"R{review_counter:05d}",
                    "category": "vendors",
                    "name": name_raw,
                    "name_display": str(name_raw).strip(),
                    "dob": "",
                    "ssn_last4": "",
                    "role": "",
                    "status": "",
                    "city": str(_row_get(row, city_col) if city_col else "").strip(),
                    "state": state,
                    "classification": classification,
                    **match_result,
                }
            )
            review_counter += 1

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
        "vendor_source_rows": int(len(df_vendor)) if df_vendor is not None else 0,
        "vendor_header_row": vendor_header_row,
        "vendor_entity_column": entity_col if df_vendor is not None else "",
        "vendor_city_column": city_col if df_vendor is not None else "",
        "vendor_state_column": state_col if df_vendor is not None else "",
    }

    if oig_path:
        metadata["oig_file_hash"] = file_sha256(oig_path)
    if sam_path:
        metadata["sam_file_hash"] = file_sha256(sam_path)

    write_audit_workbook(audit_path, config.client_name, month, results, metadata)
    write_metadata(run_dir, metadata)
    write_run_log(run_dir, "Exclusion check completed successfully.")

    # ---------------- PDF REPORTS ----------------
    staff_pdf = run_dir / "Staff_Report.pdf"
    board_pdf = run_dir / "Board_Report.pdf"
    vendor_pdf = run_dir / "Vendor_Report.pdf"

    generate_pdf_report(staff_pdf, config.client_name, month, "Staff Exclusion Report", results["staff"])
    generate_pdf_report(board_pdf, config.client_name, month, "Board Exclusion Report", results["board"])
    generate_pdf_report(vendor_pdf, config.client_name, month, "Vendor Exclusion Report", results["vendors"])

    # ---------------- COPY TO CLIENT DROPBOX FOLDER ----------------
    # Copy outputs next to the first provided source file (Dropbox folder),
    # under: ExclusionReports/
    source_file = None
    for p in (staff_path, board_path, vendor_path):
        if p and str(p).strip():
            source_file = Path(str(p)).expanduser()
            break

    copied_to = ""
    if source_file:
        try:
            base = source_file.parent
            # Prefer resolve() when possible, but do not die on weird symlinks/permissions
            try:
                base = base.resolve()
            except Exception:
                base = base.absolute()

            dropbox_output_dir = base / "ExclusionReports"
            dropbox_output_dir.mkdir(parents=True, exist_ok=True)

            files_to_copy = [audit_path, staff_pdf, board_pdf, vendor_pdf]
            for f in files_to_copy:
                if f.exists():
                    dest = dropbox_output_dir / f.name
                    shutil.copy2(f, dest)

            copied_to = str(dropbox_output_dir)
            metadata["copied_to"] = copied_to
            write_run_log(run_dir, f"Copied outputs to: {copied_to}")
            write_metadata(run_dir, metadata)
            print(f"Copied outputs to: {copied_to}")

        except Exception as e:
            metadata["copy_error"] = str(e)
            write_run_log(run_dir, f"WARNING: Copy to Dropbox folder failed: {e}")
            write_metadata(run_dir, metadata)
            print(f"WARNING: Copy to Dropbox folder failed: {e}")

    return {
        "run_directory": str(run_dir),
        "audit_file": str(audit_path),
        "staff_pdf": str(staff_pdf),
        "board_pdf": str(board_pdf),
        "vendor_pdf": str(vendor_pdf),
        "copied_to": copied_to,
    }
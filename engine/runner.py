import sqlite3
import shutil
import re
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


def _resolve_optional_column(df: pd.DataFrame, requested: str | None, fallbacks: list[str]) -> str | None:
    try:
        return _resolve_column(df, requested, fallbacks)
    except Exception:
        return None


def _infer_file_type(path, configured_file_type: str | None) -> str:
    cfg = str(configured_file_type or "auto").strip().casefold()
    if cfg in {"csv", "excel"}:
        return cfg

    ext = str(Path(path).suffix).strip().casefold()
    if ext == ".csv":
        return "csv"
    if ext in {".xlsx", ".xls", ".xlsm"}:
        return "excel"
    raise ValueError(f"Unsupported file extension for {path}: {ext}")


def _read_tabular_preview(path, file_type: str, nrows: int, delimiter: str | None):
    if file_type == "excel":
        return pd.read_excel(path, dtype=str, header=None, nrows=nrows).fillna("")

    if delimiter and str(delimiter).strip() and str(delimiter).strip().lower() != "auto":
        return pd.read_csv(path, dtype=str, header=None, nrows=nrows, sep=str(delimiter)).fillna("")

    # Auto delimiter detection for CSV if requested/unspecified.
    return pd.read_csv(path, dtype=str, header=None, nrows=nrows, sep=None, engine="python").fillna("")


def _detect_header_row(path, file_type: str, tokens: list[str], delimiter: str | None = None) -> int | None:
    """
    Detect first header row containing all expected tokens.
    """
    if not tokens:
        return None
    try:
        preview = _read_tabular_preview(path, file_type, nrows=50, delimiter=delimiter)
    except Exception:
        return None

    tokens_cf = [str(t).strip().casefold() for t in tokens if str(t).strip()]
    if not tokens_cf:
        return None

    for idx, row in preview.iterrows():
        values = [str(v or "").strip().casefold() for v in row.tolist() if str(v).strip()]
        if not values:
            continue
        if all(any(tok == val or tok in val for val in values) for tok in tokens_cf):
            return int(idx)
    return None


def _read_table_with_config(path, section: dict, category: str):
    """
    Read CSV/Excel based on config + file extension.
    Supports:
      - file_type: auto|csv|excel
      - header_row: auto|int
      - skip_rows: int
      - delimiter: ','|...|'auto' (csv only)
      - true_header_tokens: [..] (optional)
    """
    file_type = _infer_file_type(path, section.get("file_type"))
    delimiter = section.get("delimiter")
    skip_rows = int(section.get("skip_rows", 0) or 0)
    header_cfg = section.get("header_row")

    default_tokens = {
        "staff": ["first", "last", "dob"],
        "board": ["name", "dob"],
        "vendors": ["name", "city", "state"],
    }.get(category, [])
    cfg_tokens = section.get("true_header_tokens")
    header_tokens = cfg_tokens if isinstance(cfg_tokens, list) else default_tokens

    detected_header_row = None
    if isinstance(header_cfg, str) and header_cfg.strip().casefold() == "auto":
        detected_header_row = _detect_header_row(path, file_type, header_tokens, delimiter=delimiter)
    elif header_cfg is None and category == "vendors":
        # Preserve robust vendor behavior even if not explicitly configured.
        detected_header_row = _detect_header_row(path, file_type, header_tokens, delimiter=delimiter)
    elif header_cfg not in (None, ""):
        detected_header_row = int(header_cfg)

    if file_type == "excel":
        if detected_header_row is not None:
            df = pd.read_excel(path, dtype=str, header=detected_header_row).fillna("")
        elif skip_rows:
            df = pd.read_excel(path, dtype=str, skiprows=skip_rows).fillna("")
        else:
            df = pd.read_excel(path, dtype=str).fillna("")
    else:
        csv_kwargs = {"dtype": str}
        if delimiter and str(delimiter).strip():
            delim = str(delimiter).strip()
            if delim.lower() == "auto":
                csv_kwargs.update({"sep": None, "engine": "python"})
            else:
                csv_kwargs.update({"sep": delim})
        if detected_header_row is not None:
            csv_kwargs["header"] = detected_header_row
        elif skip_rows:
            csv_kwargs["skiprows"] = skip_rows
        df = pd.read_csv(path, **csv_kwargs).fillna("")

    df = _strip_columns(df)
    return df, file_type, detected_header_row, skip_rows


def _parse_city_state_zip(value: str) -> tuple[str, str, str]:
    text = str(value or "").strip()
    if not text:
        return "", "", ""

    # Find the first valid "City, ST ZIP" pattern even if cell repeats it.
    m = re.search(
        r"(?P<city>[A-Za-z0-9 .'\-/&]+?),\s*(?P<state>[A-Za-z]{2})\s*(?P<zip>\d{5}(?:-\d{4})?)?",
        text,
    )
    if m:
        city = (m.group("city") or "").strip()
        state = (m.group("state") or "").strip().upper()
        zip_code = (m.group("zip") or "").strip()
        return city, state, zip_code

    return text, "", ""


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
    staff_file_type = ""
    board_file_type = ""
    vendor_file_type = ""
    staff_header_row = None
    board_header_row = None
    vendor_header_row = None
    staff_skip_rows = 0
    board_skip_rows = 0
    vendor_skip_rows = 0

    # ---------------- STAFF ----------------
    if staff_path:
        staff_section = config.section("staff")
        df_staff, staff_file_type, staff_header_row, staff_skip_rows = _read_table_with_config(
            staff_path, staff_section, "staff"
        )

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
                    "first_name_display": str(_row_get(row, staff_section["first_name"])).strip(),
                    "middle_name_display": str(_row_get(row, staff_section.get("middle_name"))).strip(),
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
                    "address": str(_row_get(row, staff_section.get("address"))).strip(),
                    "city": str(_row_get(row, staff_section.get("city"))).strip(),
                    "state": str(_row_get(row, staff_section.get("state"))).strip(),
                    "zip_display": str(_row_get(row, staff_section.get("zip"))).strip(),
                    **match_result,
                }
            )
            review_counter += 1

    # ---------------- BOARD ----------------
    if board_path:
        board_section = config.section("board")
        df_board, board_file_type, board_header_row, board_skip_rows = _read_table_with_config(
            board_path, board_section, "board"
        )

        board_address_col = _resolve_optional_column(
            df_board,
            board_section.get("address"),
            ["Address", "Address 1", "Address1", "Street Address"],
        )
        board_city_col = _resolve_optional_column(
            df_board,
            board_section.get("city"),
            ["City", "Town", "Municipality"],
        )
        board_state_col = _resolve_optional_column(
            df_board,
            board_section.get("state"),
            ["State", "ST", "Province"],
        )
        board_zip_col = _resolve_optional_column(
            df_board,
            board_section.get("zip"),
            ["Zip", "ZIP", "Zip Code", "Postal Code"],
        )
        board_city_state_zip_col = _resolve_optional_column(
            df_board,
            board_section.get("city_state_zip"),
            ["City, State, Zip", "CITY, STATE, ZIP", "City/State/Zip", "Location"],
        )
        board_phone_col = _resolve_optional_column(
            df_board,
            board_section.get("phone"),
            ["Phone", "Phone #", "Phone Number", "Telephone", "Cell Phone"],
        )
        board_service_year_col = _resolve_optional_column(
            df_board,
            board_section.get("service_year"),
            ["Service Year", "Years of Service", "Term", "Board Service Year", "Effective Date", "Start Date"],
        )
        board_email_col = _resolve_optional_column(
            df_board,
            board_section.get("email"),
            ["Email", "Email Address", "E-mail", "E-mail Address"],
        )

        for _, row in df_board.iterrows():
            full_name = _row_get(row, board_section["name_column"])
            tokens = str(full_name).split()
            first = tokens[0] if tokens else ""
            last = tokens[-1] if len(tokens) > 1 else ""
            first, last, middle, full = normalize_person_name(first, last)

            dob_iso, dob_compact = normalize_dob(_row_get(row, board_section["dob"]))
            zip_code = normalize_zip(_row_get(row, board_section.get("zip")))
            ssn_last4 = extract_ssn_last4(_row_get(row, board_section.get("ssn")))

            city_raw = str(_row_get(row, board_city_col)).strip() if board_city_col else ""
            state_raw = str(_row_get(row, board_state_col)).strip() if board_state_col else ""
            zip_raw = str(_row_get(row, board_zip_col)).strip() if board_zip_col else ""

            # If source uses a single "City, State, Zip" column (or duplicate mapping),
            # parse it into normalized components for matching/reporting.
            same_location_col = (
                board_city_col
                and board_state_col
                and board_zip_col
                and len({board_city_col, board_state_col, board_zip_col}) == 1
            )
            combo_col = board_city_state_zip_col or (board_city_col if same_location_col else None)
            if combo_col:
                if same_location_col:
                    # Prevent the same raw text from being copied to state/zip on fallback.
                    state_raw = ""
                    zip_raw = ""
                combo_city, combo_state, combo_zip = _parse_city_state_zip(_row_get(row, combo_col))
                city_raw = combo_city or city_raw
                state_raw = combo_state or state_raw
                zip_raw = combo_zip or zip_raw

            match_result = match_person(
                conn,
                first,
                last,
                dob_compact=dob_compact,
                zip_code=normalize_zip(zip_raw) or zip_code,
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
                    "address": str(_row_get(row, board_address_col)).strip() if board_address_col else "",
                    "city": city_raw,
                    "state": state_raw,
                    "zip_display": zip_raw,
                    "phone": str(_row_get(row, board_phone_col)).strip() if board_phone_col else "",
                    "service_year": str(_row_get(row, board_service_year_col)).strip() if board_service_year_col else "",
                    "email": str(_row_get(row, board_email_col)).strip() if board_email_col else "",
                    **match_result,
                }
            )
            review_counter += 1

    # ---------------- VENDORS ----------------
    df_vendor = None
    entity_col = ""
    city_col = None
    state_col = None
    taxid_col = None
    address_col = None
    address2_col = None
    vendor_id_col = None
    if vendor_path:
        vendor_section = config.section("vendors")
        df_vendor, vendor_file_type, vendor_header_row, vendor_skip_rows = _read_table_with_config(
            vendor_path, vendor_section, "vendors"
        )

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
            address_col = _resolve_column(
                df_vendor,
                vendor_section.get("address"),
                fallbacks=["Address", "ADDRESS", "Address 1", "ADDRESS1", "Street", "Street Address"],
            )
        except Exception:
            address_col = None

        try:
            address2_col = _resolve_column(
                df_vendor,
                vendor_section.get("address2"),
                fallbacks=["Address 2", "ADDRESS2", "Address2", "Address line 2", "Addr2"],
            )
        except Exception:
            address2_col = None

        try:
            vendor_id_col = _resolve_column(
                df_vendor,
                vendor_section.get("vendor_id"),
                fallbacks=["Vendor ID", "VendorID", "ID", "Supplier ID", "Payee ID"],
            )
        except Exception:
            vendor_id_col = None

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
                    "address": str(_row_get(row, address_col) if address_col else "").strip(),
                    "address2": str(_row_get(row, address2_col) if address2_col else "").strip(),
                    "zip_display": str(_row_get(row, zip_col) if zip_col else "").strip(),
                    "vendor_id": str(_row_get(row, vendor_id_col) if vendor_id_col else "").strip(),
                    "tax_id_display": str(tax_id).strip(),
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
        "staff_file_type": staff_file_type,
        "board_file_type": board_file_type,
        "vendor_file_type": vendor_file_type,
        "staff_header_row": staff_header_row,
        "board_header_row": board_header_row,
        "vendor_source_rows": int(len(df_vendor)) if df_vendor is not None else 0,
        "vendor_header_row": vendor_header_row,
        "staff_skip_rows": staff_skip_rows,
        "board_skip_rows": board_skip_rows,
        "vendor_skip_rows": vendor_skip_rows,
        "vendor_entity_column": entity_col if df_vendor is not None else "",
        "vendor_city_column": city_col if df_vendor is not None else "",
        "vendor_state_column": state_col if df_vendor is not None else "",
        "vendor_address_column": address_col if df_vendor is not None else "",
        "vendor_address2_column": address2_col if df_vendor is not None else "",
        "vendor_id_column": vendor_id_col if df_vendor is not None else "",
        "vendor_tax_id_column": taxid_col if df_vendor is not None else "",
        "staff_review_required": sum(1 for row in results["staff"] if row.get("review_required")),
        "board_review_required": sum(1 for row in results["board"] if row.get("review_required")),
        "vendor_review_required": sum(1 for row in results["vendors"] if row.get("review_required")),
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
    generate_pdf_report(board_pdf, config.client_name, month, "Board Members Exclusion Report", results["board"])
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
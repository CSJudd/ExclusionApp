import sqlite3
from rapidfuzz import fuzz

from engine.normalizer import normalize_person_name


FUZZ_STRONG = 95
FUZZ_POSSIBLE = 90


def match_person(conn, first, last, dob_compact=None, city=None, state=None, zip_code=None):
    """
    Returns:
    {
        "oig_status": "...",
        "oig_date": "...",
        "sam_status": "...",
        "sam_date": "...",
        "reason": "..."
    }
    """

    result = {
        "oig_status": "NOT FOUND",
        "oig_date": "",
        "sam_status": "NOT FOUND",
        "sam_date": "",
        "reason": ""
    }

    cur = conn.cursor()

    # --- OIG PERSON MATCH ---
    cur.execute("""
        SELECT first, last, dob_compact, exclusion_date
        FROM oig_people
        WHERE last = ?
    """, (last,))

    rows = cur.fetchall()

    for row in rows:
        db_first, db_last, db_dob, exclusion_date = row

        score = fuzz.ratio(first, db_first)

        # Exact + DOB match
        if first == db_first and dob_compact and dob_compact == db_dob:
            result["oig_status"] = "CONFIRMED"
            result["oig_date"] = exclusion_date
            result["reason"] = "Exact first+last+DOB"
            break

        # Strong fuzzy + DOB
        if score >= FUZZ_STRONG and dob_compact and dob_compact == db_dob:
            result["oig_status"] = "CONFIRMED"
            result["oig_date"] = exclusion_date
            result["reason"] = f"Fuzzy first ({score}) + DOB"
            break

        # Possible
        if score >= FUZZ_POSSIBLE:
            result["oig_status"] = "POSSIBLE"
            result["oig_date"] = exclusion_date
            result["reason"] = f"Fuzzy first ({score})"
            break

    # --- SAM PERSON MATCH ---
    cur.execute("""
        SELECT first, last, exclusion_date, city, state, zip
        FROM sam_people
        WHERE last = ?
    """, (last,))

    rows = cur.fetchall()

    for row in rows:
        db_first, db_last, exclusion_date, db_city, db_state, db_zip = row
        score = fuzz.ratio(first, db_first)

        if first == db_first:
            # Secondary signal check
            if zip_code and db_zip and zip_code == db_zip:
                result["sam_status"] = "CONFIRMED"
                result["sam_date"] = exclusion_date
                break

            if city and db_city and city.upper() == db_city:
                result["sam_status"] = "CONFIRMED"
                result["sam_date"] = exclusion_date
                break

        if score >= FUZZ_STRONG:
            result["sam_status"] = "POSSIBLE"
            result["sam_date"] = exclusion_date
            break

    return result
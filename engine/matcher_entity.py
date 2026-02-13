import sqlite3
from rapidfuzz import fuzz

FUZZ_STRONG = 95


def _set_review(result, *, source, candidate_name, exclusion_date, note, needed_data=""):
    if result.get("review_required"):
        return
    result["review_required"] = True
    result["review_source"] = source
    result["review_candidate_name"] = candidate_name
    result["review_candidate_exclusion_date"] = exclusion_date or ""
    result["review_note"] = note
    result["review_needed_data"] = needed_data


def match_entity(conn, entity_name, city=None, state=None, zip_code=None):
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
        "reason": "",
        "review_required": False,
        "review_source": "",
        "review_candidate_name": "",
        "review_candidate_exclusion_date": "",
        "review_note": "",
        "review_needed_data": "",
    }

    cur = conn.cursor()

    # --- OIG ENTITY MATCH ---
    cur.execute("""
        SELECT name, exclusion_date
        FROM oig_entities
        WHERE name = ?
    """, (entity_name,))

    exact_oig = cur.fetchone()

    if exact_oig:
        result["oig_status"] = "CONFIRMED"
        result["oig_date"] = exact_oig[1]
        result["reason"] = "Exact entity name match (OIG)"
    else:
        # Fuzzy OIG
        cur.execute("SELECT name, exclusion_date FROM oig_entities")
        for db_name, exclusion_date in cur.fetchall():
            score = fuzz.ratio(entity_name, db_name)
            if score >= FUZZ_STRONG:
                _set_review(
                    result,
                    source="OIG Entities",
                    candidate_name=db_name,
                    exclusion_date=exclusion_date,
                    note=f"High-similarity OIG entity name match (score={score}).",
                    needed_data="Tax ID / address corroboration",
                )
                break

    # --- SAM ENTITY MATCH ---
    cur.execute("""
        SELECT name, exclusion_date, city, state, zip
        FROM sam_entities
        WHERE name = ?
    """, (entity_name,))

    exact_sam = cur.fetchone()

    if exact_sam:
        result["sam_status"] = "CONFIRMED"
        result["sam_date"] = exact_sam[1]
    else:
        # Fuzzy SAM
        cur.execute("""
            SELECT name, exclusion_date, city, state, zip
            FROM sam_entities
        """)
        for db_name, exclusion_date, db_city, db_state, db_zip in cur.fetchall():
            score = fuzz.ratio(entity_name, db_name)
            if score >= FUZZ_STRONG:
                # Secondary signal check
                if state and db_state and state.upper() == db_state:
                    _set_review(
                        result,
                        source="SAM Entities",
                        candidate_name=db_name,
                        exclusion_date=exclusion_date,
                        note=f"High-similarity SAM entity match (score={score}) with state corroboration.",
                        needed_data="Tax ID / exact legal name confirmation",
                    )
                    break
                if zip_code and db_zip and zip_code == db_zip:
                    _set_review(
                        result,
                        source="SAM Entities",
                        candidate_name=db_name,
                        exclusion_date=exclusion_date,
                        note=f"High-similarity SAM entity match (score={score}) with zip corroboration.",
                        needed_data="Tax ID / exact legal name confirmation",
                    )
                    break

    return result
import sqlite3
from rapidfuzz import fuzz

FUZZ_STRONG = 95


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
        "reason": ""
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
                result["oig_status"] = "POSSIBLE"
                result["oig_date"] = exclusion_date
                result["reason"] = f"Fuzzy entity match (OIG) score={score}"
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
                    result["sam_status"] = "POSSIBLE"
                    result["sam_date"] = exclusion_date
                    break
                if zip_code and db_zip and zip_code == db_zip:
                    result["sam_status"] = "POSSIBLE"
                    result["sam_date"] = exclusion_date
                    break

    return result
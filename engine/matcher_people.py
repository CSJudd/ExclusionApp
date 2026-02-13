import sqlite3
from rapidfuzz import fuzz

from engine.normalizer import normalize_person_name


FUZZ_STRONG = 95
FUZZ_POSSIBLE = 90


def _set_review(result, *, source, candidate_name, exclusion_date, note, needed_data=""):
    if result.get("review_required"):
        return
    result["review_required"] = True
    result["review_source"] = source
    result["review_candidate_name"] = candidate_name
    result["review_candidate_exclusion_date"] = exclusion_date or ""
    result["review_note"] = note
    result["review_needed_data"] = needed_data


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
        "reason": "",
        "review_required": False,
        "review_source": "",
        "review_candidate_name": "",
        "review_candidate_exclusion_date": "",
        "review_note": "",
        "review_needed_data": "",
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
            candidate_name = f"{db_first} {db_last}".strip()
            if not dob_compact:
                _set_review(
                    result,
                    source="OIG People",
                    candidate_name=candidate_name,
                    exclusion_date=exclusion_date,
                    note=f"High-similarity OIG name match (score={score}) but DOB missing in source record.",
                    needed_data="DOB",
                )
            elif db_dob and dob_compact != db_dob:
                _set_review(
                    result,
                    source="OIG People",
                    candidate_name=candidate_name,
                    exclusion_date=exclusion_date,
                    note=f"High-similarity OIG name match (score={score}) but DOB does not match reference.",
                    needed_data="Confirm DOB / SSN last4",
                )

    # --- SAM PERSON MATCH (STRICT POLICY) ---
    # SAM individual records frequently lack deterministic identifiers (e.g., DOB/SSN),
    # so only corroborated location matches are treated as actionable.
    # Policy:
    # - CONFIRMED: exact first+last and (zip match OR city+state+zip match)
    # - all other SAM-person cases: NOT FOUND (no review item)
    cur.execute("""
        SELECT first, last, exclusion_date, city, state, zip
        FROM sam_people
        WHERE last = ?
    """, (last,))

    rows = cur.fetchall()

    for row in rows:
        db_first, db_last, exclusion_date, db_city, db_state, db_zip = row
        if first == db_first:
            state_matches = bool(state and db_state and state.upper() == db_state)
            city_matches = bool(city and db_city and city.upper() == db_city)
            zip_matches = bool(zip_code and db_zip and zip_code == db_zip)

            # Strong corroboration path:
            # - zip match, and
            # - if city/state are present in source, they must align too
            if zip_matches and ((not city and not state) or (city_matches and state_matches)):
                result["sam_status"] = "CONFIRMED"
                result["sam_date"] = exclusion_date
                break
        # Intentionally ignore SAM person name-only/fuzzy/non-corroborated cases.

    return result
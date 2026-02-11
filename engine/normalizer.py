# engine/normalizer.py

import re
from datetime import datetime

BUSINESS_SUFFIXES = {
    "LLC", "L.L.C.", "INC", "INC.", "CORP", "CORPORATION",
    "CO", "COMPANY", "LTD", "LIMITED", "PLLC", "PC",
    "LP", "LLP", "ASSOCIATES", "GROUP", "SERVICES"
}

PERSON_SUFFIXES = {"JR", "SR", "III", "IV", "II"}

def normalize_whitespace(value: str) -> str:
    return re.sub(r"\s+", " ", value.strip())

def normalize_name(value: str) -> str:
    if not value:
        return ""
    value = value.upper()
    value = re.sub(r"[^\w\s]", "", value)
    value = normalize_whitespace(value)
    return value

def remove_person_suffixes(name: str) -> str:
    tokens = name.split()
    tokens = [t for t in tokens if t not in PERSON_SUFFIXES]
    return " ".join(tokens)

def normalize_person_name(first: str, last: str, middle: str = None):
    first = normalize_name(first)
    last = normalize_name(last)
    middle = normalize_name(middle) if middle else ""
    first = remove_person_suffixes(first)
    last = remove_person_suffixes(last)
    full = normalize_whitespace(f"{first} {middle} {last}")
    return first, last, middle, full

def normalize_entity_name(name: str) -> str:
    if not name:
        return ""
    name = normalize_name(name)
    tokens = name.split()
    tokens = [t for t in tokens if t not in BUSINESS_SUFFIXES]
    return " ".join(tokens)

def normalize_zip(zip_code: str) -> str:
    if not zip_code:
        return ""
    zip_code = re.sub(r"[^\d]", "", zip_code)
    return zip_code[:5]

def normalize_dob(dob_value):
    if not dob_value:
        return None, None
    try:
        parsed = datetime.strptime(str(dob_value).strip(), "%m/%d/%Y")
    except ValueError:
        try:
            parsed = datetime.strptime(str(dob_value).strip(), "%Y-%m-%d")
        except ValueError:
            return None, None
    iso = parsed.strftime("%Y-%m-%d")
    compact = parsed.strftime("%Y%m%d")
    return iso, compact

def extract_ssn_last4(ssn_value: str):
    if not ssn_value:
        return None
    digits = re.sub(r"[^\d]", "", ssn_value)
    if len(digits) >= 4:
        return digits[-4:]
    return None

def is_ein(tax_id: str):
    if not tax_id:
        return False
    digits = re.sub(r"[^\d]", "", tax_id)
    return len(digits) == 9 and "-" in tax_id

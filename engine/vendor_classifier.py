import re

BUSINESS_SUFFIXES = {
    "LLC", "L.L.C.", "INC", "INC.", "CORP", "CORPORATION",
    "CO", "COMPANY", "LTD", "LIMITED", "PLLC", "PC",
    "LP", "LLP"
}

BUSINESS_KEYWORDS = {
    "GROUP", "SERVICES", "ASSOCIATES", "ENTERPRISES",
    "HOLDINGS", "SOLUTIONS", "CLINIC", "MEDICAL",
    "HEALTH", "THERAPY", "SUPPLY"
}


def classify_vendor(name: str, tax_id: str = None):
    """
    Returns:
        "ENTITY"
        "PERSON_VENDOR"
        "AMBIGUOUS"
    """

    if not name:
        return "AMBIGUOUS"

    upper_name = name.upper().strip()

    tokens = re.sub(r"[^\w\s]", "", upper_name).split()

    # EIN pattern check (##-####### or 9 digits)
    if tax_id:
        digits = re.sub(r"[^\d]", "", tax_id)
        if len(digits) == 9:
            return "ENTITY"

    # Explicit business suffix check
    if any(token in BUSINESS_SUFFIXES for token in tokens):
        return "ENTITY"

    # Business keyword signal
    if any(token in BUSINESS_KEYWORDS for token in tokens):
        return "ENTITY"

    # Person-like structure: two or three tokens
    if 2 <= len(tokens) <= 3:
        # Avoid classifying obvious commercial patterns as person
        if not any(token in BUSINESS_KEYWORDS for token in tokens):
            return "PERSON_VENDOR"

    # Default fallback
    return "AMBIGUOUS"
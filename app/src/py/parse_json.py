import sys
import json

def read_stdin_json():
    """Reads JSON data from standard input."""
    raw = sys.stdin.read()
    return json.loads(raw)

def split_data(data):
    form = data["form"]
    case = data.get("case", {})
    lawyer = data.get("lawyer", {})
    return form, case, lawyer

def get_matter(data):
    """ Returns the matter ID from the JSON data. """
    form, _, _ = split_data(data)
    matter = form.get("matterId", "")
    if not matter:
        print("No matter ID provided.")
        raise ValueError("Matter ID is required but was not provided.")
    return matter

def load_consultation_fields(data):
    form, _, _ = split_data(data)

    # Convert values from frontend to what the PCLaw form expects
    name_parts = form["clientName"].strip().split()
    first_name = name_parts[0] if name_parts else ""
    last_name = name_parts[-1] if len(name_parts) > 1 else ""
    middle_name = " ".join(name_parts[1:-1]) if len(name_parts) > 2 else ""

    # Determine language prefix
    lang = get_language(data)
    if lang.startswith("en"):
        desc_prefix = "Consultation in"
    else:
        desc_prefix = "Consultation en"

    description = f"{desc_prefix} {get_case_details(data).lower()}".strip()

    if form.get("isRefBarreau", False):
        default_rate = "J"
    elif form.get("isFirstConsultation", False):
        default_rate = "I"
    else:
        default_rate = "A"

    # Determine Responsible Lawyer: if not MM, DH, or TG, use "JR"
    lawyer_id = form.get("lawyerId", "")
    if lawyer_id not in ("MM", "DH", "TG"):
        responsible_lawyer = "JR"
    else:
        responsible_lawyer = lawyer_id

    # Determine "Type of Law" field logic
    if form.get("isFirstConsultation", False):
        type_of_law = "cons"
    else:
        case_type = form.get("caseType", "").lower()
        if case_type == "estate":
            type_of_law = "est"
        elif case_type == "divorce":
            type_of_law = "mat"
        else:
            type_of_law = ""

    fields = [
        (2, type_of_law), # Type of Law
        (3, default_rate), # Default Rate
        (5, responsible_lawyer), # Responsible Lawyer
        (6, description), # Description
        (7, form.get("clientTitle", "")), # Title
        (8, first_name), # First
        (9, middle_name), # Middle
        (10, last_name), # Last
        (25, form.get("clientPhone", "")), # Cell
        (29, form.get("clientEmail", "")), # E-mail 1
    ]

    return fields

def get_language(data):
    """ Returns the client language from the JSON data. """
    form, _, _ = split_data(data)
    return form.get("clientLanguage", "English").lower()

def get_case_details(data):
    """
    Returns a string with details for the current case type,
    in English or French depending on client language.
    """
    form, case_details, _ = split_data(data)
    
    lang = form.get("clientLanguage", "English").lower()
    case_type = form.get("caseType", "").lower()
    
    if not case_type:
        return ""
    def get(key):
        return case_details.get(key, "")

    if lang.startswith("fr"):
        mapping = {
            "divorce": "Divorce / Droit de la famille",
            "estate": "Droit des successions",
            "employment": "Droit du travail",
            "contract": "Droit des contrats",
            "defamations": "Diffamations",
            "real_estate": "Immobilier",
            "name_change": "Changement de nom",
            "adoptions": "Adoption",
            "mandates": "Régime de protection",
            "business": "Droit des affaires",
            "assermentation": "Assermentation",
            "common": str(get("commonField")),
        }
    else:
        mapping = {
            "divorce": "Divorce / Family Law",
            "estate": "Estate Law",
            "employment": "Employment Law",
            "contract": "Contract Law",
            "defamations": "Defamations",
            "real_estate": "Real Estate",
            "name_change": "Name Change",
            "adoptions": "Adoption",
            "mandates": "Protection mandates",
            "business": "Business Law",
            "assermentation": "Assermentation",
            "common": str(get("commonField")),
        }
    
    return mapping.get(case_type, "")

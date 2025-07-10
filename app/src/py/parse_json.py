import json
import os

APP_DIR = os.path.abspath(os.path.join(__file__, "../../.."))
JSON_PATH = os.path.join(APP_DIR, "data.json")

def load_json(json_path=JSON_PATH):
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    form = data["form"]
    case = data.get("caseDetails", {})
    lawyer = data.get("lawyer", {})

    return form, case, lawyer

def load_consultation_fields():
    form, _, _ = load_json()

    # Convert values from frontend to what the PCLaw form expects
    name_parts = form["clientName"].strip().split()
    first_name = name_parts[0] if name_parts else ""
    last_name = name_parts[-1] if len(name_parts) > 1 else ""
    middle_name = " ".join(name_parts[1:-1]) if len(name_parts) > 2 else ""

    # Determine language prefix
    lang = get_language()
    if lang.startswith("en"):
        desc_prefix = "Consultation in"
    else:
        desc_prefix = "Consultation en"

    description = f"{desc_prefix} {get_case_details().lower()}".strip()

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

    fields = {
        "Type of Law": type_of_law,
        "Description": description,
        "Responsible Lawyer": responsible_lawyer,
        "Default Rate": default_rate,
        "Title": form.get("clientTitle", ""),
        "First": first_name,
        "Middle": middle_name,
        "Last": last_name,
        "Cell": form.get("clientPhone", ""),
        "E-mail 1": form.get("clientEmail", ""),
    }

    return fields

def get_language():
    """ Returns the client language from the JSON data.
    Defaults to "english" if not found.
    """
    form, _, _ = load_json()
    return form.get("clientLanguage", "English").lower()

def get_case_details():
    """
    Returns a string with details for the current case type,
    in English or French depending on client language.
    """
    form, case_details, _ = load_json()
    case_type = form.get("caseType", "").lower()
    lang = get_language()
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
            "estate": "Successions / Estate Law",
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

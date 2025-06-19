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
    form, case, lawyer = load_json()

    # Convert values from frontend to what the PCLaw form expects
    full_name = form["clientName"].strip().split(" ", 1)
    first_name = full_name[0]
    last_name = full_name[1] if len(full_name) > 1 else ""

    # Determine language prefix
    lang = get_language()
    if lang.startswith("en"):
        desc_prefix = "Consultation in"
    else:
        desc_prefix = "Consultation en"

    case_type = form.get("caseType", "").lower()
    if case_type == "common":
        description = f"{desc_prefix} {case.get('commonField', '').lower()}".strip()
    else:
        description = f"{desc_prefix} {case_type}".strip()

    if form.get("isRefBarreau", False):
        default_rate = "J"
    elif form.get("isFirstConsultation", False):
        default_rate = "I"
    else:
        default_rate = "A"

    fields = {
        "Type of Law": "cons",
        "Description": description,
        "Responsible Lawyer": form.get("lawyerId", ""),
        "Default Rate": default_rate,
        "Title": form.get("clientTitle", ""),
        "First": first_name,
        "Last": last_name,
        "Cell": form.get("clientPhone", ""),
        "E-mail 1": form.get("clientEmail", ""),
    }

    return fields

def get_language():
    """ Returns the client language from the JSON data.
    Defaults to "English" if not found.
    """
    form, _, _ = load_json()
    return form.get("clientLanguage", "English").lower()

def get_case_details():
    """
    Returns an HTML string with details for the current case type and details from the JSON form.
    """
    form, case_details, _ = load_json()
    case_type = form.get("caseType", "").lower()
    if not case_type:
        return ""
    def get(key):
        return case_details.get(key, "")
    if case_type == "divorce":
        spouse = get("spouseName")
        conflict = get("conflictSearchDone")
        conflict_str = "✔️" if conflict is True else "❌"
        return (
            "Divorce / Family Law<br><br>"
            f"<strong>Spouse Name:</strong> {spouse}<br>"
            f"Conflict Search Done? {conflict_str}"
        )
    elif case_type == "estate":
        deceased = get("deceasedName")
        executor = get("executorName")
        conflict = get("conflictSearchDone")
        conflict_str = "✔️" if conflict is True else "❌"
        return (
            "Successions / Estate Law<br><br>"
            f"<strong>Deceased Name:</strong> {deceased}<br>"
            f"<strong>Executor Name:</strong> {executor}<br>"
            f"Conflict Search Done? {conflict_str}"
        )
    elif case_type == "employment":
        employer = get("employerName")
        return (
            "Employment Law<br><br>"
            f"<strong>Employer Name:</strong> {employer}"
        )
    elif case_type == "contract":
        other_party = get("otherPartyName")
        return (
            "Contract Law<br><br>"
            f"<strong>Other Party:</strong> {other_party}"
        )
    elif case_type == "defamations":
        return "Defamations"
    elif case_type == "real_estate":
        return "Real Estate"
    elif case_type == "name_change":
        return "Name Change"
    elif case_type == "adoptions":
        return "Adoptions"
    elif case_type == "mandates":
        mandate_details = get("mandateDetails")
        return (
            "Regimes de Protection / Mandates<br><br>"
            f"<strong>Mandate Details:</strong> {mandate_details}"
        )
    elif case_type == "business":
        business_name = get("businessName")
        return (
            "Business Law<br><br>"
            f"<strong>Business Name:</strong> {business_name}"
        )
    elif case_type == "assermentation":
        return "Assermentation"
    elif case_type == "common":
        return str(get("commonField"))
    else:
        return ""


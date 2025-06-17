import json
import os

root_dir = os.path.abspath(os.path.join(__file__, "../../../"))
json_path = os.path.join(root_dir, "app", "data.json")

def load_consultation_fields(json_path=json_path):
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    form = data["form"]
    case = data.get("caseDetails", {})
    lawyer = data.get("lawyer", {})

    # Convert values from frontend to what the PCLaw form expects
    full_name = form["clientName"].strip().split(" ", 1)
    first_name = full_name[0]
    last_name = full_name[1] if len(full_name) > 1 else ""

    # Determine language prefix
    lang = form.get("clientLanguage", "English").lower()
    if lang == "English":
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

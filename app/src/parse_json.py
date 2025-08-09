from config import *
import json

def read_json(file_path=None):
    """Reads JSON data from a file or standard input."""
    try:
        if file_path:
            with open(file_path, 'r') as f:
                data = json.load(f)
        else:
            raw = sys.stdin.read()
            data = json.loads(raw)
        return data
    except FileNotFoundError:
        alert_error(f"File not found: {file_path}")
        raise FileNotFoundError(f"File not found: {file_path}")
    except json.JSONDecodeError as e:
        alert_error(f"Invalid JSON: {e.msg}")
        raise ValueError(f"Invalid JSON: {e.msg}") from e
    except Exception as e:
        alert_error(f"Error reading JSON: {str(e)}")
        raise ValueError(f"Error reading JSON: {str(e)}") from e

def split_data(data):
    """Splits JSON data into form, case, and lawyer sections."""
    try:
        form = data["form"]
        case = data["case"]
        lawyer = data["lawyer"]
        return form, case, lawyer
    except KeyError as e:
        alert_error(f"Missing key in JSON data: {str(e)}")
        raise ValueError(f"Missing key in JSON data: {str(e)}") from e

def get_matter(data):
    """ Returns the matter ID from the JSON data. """
    try:
        form, _, _ = split_data(data)
        matter = form.get("matterId", "")
        if not matter:
            print("No matter ID provided.")
            raise ValueError("Matter ID is required but was not provided.")
        return matter
    except KeyError as e:
        alert_error(f"Missing key in JSON data: {str(e)}")
        raise ValueError(f"Missing key in JSON data: {str(e)}") from e
    except Exception as e:
        alert_error(f"Error retrieving matter ID: {str(e)}")
        raise ValueError(f"Error retrieving matter ID: {str(e)}") from e

def load_consultation_fields(data):
    try:
        form, case_details, _ = split_data(data)

        # Extract commonly used values once
        client_name = form["clientName"].strip()
        client_language = form.get("clientLanguage", "English")
        case_type = form.get("caseType", "")
        lawyer_id = form.get("lawyerId", "")

        # Convert values from frontend to what the PCLaw form expects
        name_parts = client_name.split()
        first_name = name_parts[0] if name_parts else ""
        last_name = name_parts[-1] if len(name_parts) > 1 else ""
        middle_name = " ".join(name_parts[1:-1]) if len(name_parts) > 2 else ""

        # Determine language prefix for description
        lang = client_language.lower()
        if lang.startswith("en"):
            desc_prefix = "Consultation in"
        else:
            desc_prefix = "Consultation en"

        # Use our unified get_case_details function - no redundant parsing!
        case_details_text = get_case_details(
            case_type=case_type,
            case_details_dict=case_details,
            client_language=client_language,
            format_type="name_only",
            include_details=False
        )

        description = f"{desc_prefix} {case_details_text.lower()}".strip()

        if form.get("isRefBarreau", False):
            default_rate = "J"
        elif form.get("isFirstConsultation", False):
            default_rate = "I"
        else:
            default_rate = "A"

        # Determine Responsible Lawyer: if not MM, DH, or TG, use "JR"
        if lawyer_id not in ("MM", "DH", "TG"):
            responsible_lawyer = "JR"
        else:
            responsible_lawyer = lawyer_id

        law_types = {
            "divorce": "mat",
            "estate": "est", 
            "employment": "Lab",
            "contract": "con",
            "defamations": "def",
            "real_estate": "re",
            "name_change": "nam",
            "adoptions": "ado",
            "mandates": "Rem",
            "business": "Com",
            "assermentation": "oa",
            "common" : ""
        }
        # Determine "Type of Law" field logic
        if form.get("isFirstConsultation", False):
            type_of_law = "cons"
        else:
            case_type = form.get("caseType", "").lower()
            type_of_law = law_types.get(case_type, "")

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
    except KeyError as e:
        alert_error(f"Missing key in JSON data: {str(e)}")
        raise ValueError(f"Missing key in JSON data: {str(e)}") from e
    except Exception as e:
        alert_error(f"Error loading consultation fields: {str(e)}")
        raise ValueError(f"Error loading consultation fields: {str(e)}") from e

def get_language(data):
    """ Returns the client language from the JSON data. """
    form, _, _ = split_data(data)
    return form.get("clientLanguage", "English").lower()

def get_case_details(case_type, case_details_dict, client_language="English", format_type="name_only", include_details=False):
    """
    Returns case details in various formats based on provided parameters.
    
    Args:
        case_type: String representing the type of case (e.g., "divorce", "estate")
        case_details_dict: Dictionary containing case-specific details
        client_language: Language preference for localization
        format_type: "name_only", "detailed_text", or "detailed_html"
        include_details: Whether to include specific case field details
    """
    if not case_type:
        return ""
    
    def get_detail(key):
        """Helper to safely extract details from the case details dictionary"""
        return case_details_dict.get(key, "") if case_details_dict else ""
    
    def conflict_check(done):
        """Helper to convert boolean conflict check to visual indicator"""
        return "✔️" if done else "❌"
    
    case_type = case_type.lower()
    lang = client_language.lower()
    
    # Define case names in both languages - this centralizes all localization
    if lang.startswith("fr"):
        case_names = {
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
            "common": str(get_detail("commonField")),
        }
    else:
        case_names = {
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
            "common": str(get_detail("commonField")),
        }
    
    base_name = case_names.get(case_type, "")
    
    # If only the name is requested, return it immediately
    if format_type == "name_only" or not include_details:
        return base_name
    
    # Build detailed information based on case type
    if not case_details_dict:
        return base_name
        
    # Set up formatting based on output type
    separator = "\n\n" if format_type == "detailed_text" else "<br><br>"
    field_separator = "\n" if format_type == "detailed_text" else "<br>"
    bold_open = "" if format_type == "detailed_text" else "<strong>"
    bold_close = "" if format_type == "detailed_text" else "</strong>"
    
    details = base_name
    
    # Add case-specific details based on case type
    if case_type == "divorce":
        conflict_done = case_details_dict.get("conflictSearchDone", False)
        details += f"{separator}{bold_open}Spouse Name:{bold_close} {get_detail('spouseName')}{field_separator}"
        details += f"Conflict Search Done? {conflict_check(conflict_done)}"
        
    elif case_type == "estate":
        conflict_done = case_details_dict.get("conflictSearchDone", False)
        details += f"{separator}{bold_open}Deceased Name:{bold_close} {get_detail('deceasedName')}{field_separator}"
        details += f"{bold_open}Executor Name:{bold_close} {get_detail('executorName')}{field_separator}"
        details += f"Conflict Search Done? {conflict_check(conflict_done)}"
        
    elif case_type == "employment":
        details += f"{separator}{bold_open}Employer Name:{bold_close} {get_detail('employerName')}"
        
    elif case_type == "contract":
        details += f"{separator}{bold_open}Other Party:{bold_close} {get_detail('otherPartyName')}"
        
    elif case_type == "mandates":
        details += f"{separator}{bold_open}Mandate Details:{bold_close} {get_detail('mandateDetails')}"
        
    elif case_type == "business":
        details += f"{separator}{bold_open}Business Name:{bold_close} {get_detail('businessName')}"
    
    return details

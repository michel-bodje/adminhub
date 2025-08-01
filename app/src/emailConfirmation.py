from office_utils import *
from parse_json import *
import win32com.client as COM
from datetime import datetime

def process_email_confirmation(data):
    """
    Process email confirmation with the provided data.
    
    Args:
        data (dict): The parsed JSON data containing form information
        
    Returns:
        dict: Result of the email confirmation process
    """
    try:
        # Use the passed data instead of reading from stdin
        form, _, lawyer = split_data(data)

        client_email = form["clientEmail"]
        client_language = form.get("clientLanguage", "English").lower()
        location = form.get("location", "")
        is_ref_barreau = form.get("isRefBarreau", False)
        is_first_consult = form.get("isFirstConsultation", False)
        appointment_date = form.get("appointmentDate", "")
        appointment_time = form.get("appointmentTime", "")
        lawyer_name = lawyer.get("name", "")
        lawyer_id = lawyer.get("id", "")

        lang = "fr" if client_language.startswith("fr") else "en"
        template_name = location.capitalize() + ".html"
        template_path = os.path.join(TEMPLATES_DIR, lang, template_name)

        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template not found: {template_path}")

        with open(template_path, "r", encoding="utf-8") as f:
            html_body = f.read()

        # Handle appointment date and time
        if appointment_date and appointment_time:
            slot = datetime.strptime(f"{appointment_date} {appointment_time}", "%Y-%m-%d %H:%M")

            formatted_date = format_date(slot, lang)
            formatted_time = format_time(slot, lang)

            base_rate = 60 if is_ref_barreau else 125 if is_first_consult else 350
            total_rate = add_taxes(base_rate)

            html_body = (html_body
                         .replace("{{date}}", formatted_date)
                         .replace("{{time}}", formatted_time)
                         .replace("{{rates}}", str(base_rate))
                         .replace("{{totalRates}}", f"{total_rate:.2f}"))

        lawyer_string = get_lawyer_string(lawyer_name, lawyer_id)
        html_body = html_body.replace("{{lawyerName}}", lawyer_string)

        outlook = COM.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = client_email
        mail.Subject = ("Confirmation de rendez-vous - Allen Madelin"
                        if lang == "fr" else "Confirmation of appointment - Allen Madelin")
        mail.HTMLBody = html_body
        mail.Display()

        focus_office_window(mail)
        
        # Return success result
        return {
            "status": "success",
            "message": "Confirmation email draft opened in Outlook.",
            "recipient": client_email,
            "language": lang
        }

    except Exception as e:
        alert_error(f"Error: {e}")
        raise

# Backward compatibility
def main():
    try:
        data = read_json()
        result = process_email_confirmation(data)
        if result.get("error"):
            print(f"Error: {result['error']}")
        else:
            print(result.get("message", "Email confirmation processed"))
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()

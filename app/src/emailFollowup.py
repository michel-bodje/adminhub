from office_utils import *
from parse_json import *
import win32com.client as COM
import os

def process_email_followup(data):
    try:
        form, _, _ = split_data(data)

        client_email = form["clientEmail"]
        client_language = get_language(data)
        lang = "fr" if client_language.startswith("fr") else "en"

        template_path = os.path.join(TEMPLATES_DIR, lang, "Suivi.html")
        with open(template_path, "r", encoding="utf-8") as f:
            html_body = f.read()

        outlook = COM.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = client_email
        mail.Subject = "Suivi de dossier - Allen Madelin" if lang == "fr" else "Follow-up - Allen Madelin"
        mail.HTMLBody = html_body
        mail.Display()

        focus_office_window(mail)

        # Return success result
        return {
            "status": "success",
            "message": "Follow-up email draft opened in Outlook.",
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
        result = process_email_followup(data)
        if result.get("error"):
            print(f"Error: {result['error']}")
        else:
            print(result.get("message", "Email followup processed"))
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()

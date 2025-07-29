from office_utils import *
from parse_json import *
import win32com.client as COM
import os

def draft_followup():
    try:
        data = read_json("test.json")
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
        print("Follow-up email draft opened in Outlook.")

    except Exception as e:
        alert_error(f"Error: {e}")
        raise

if __name__ == "__main__":
    draft_followup()

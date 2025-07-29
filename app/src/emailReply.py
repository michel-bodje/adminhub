from office_utils import *
from parse_json import *
import win32com.client as COM

def draft_reply():
    try:
        data = read_json()
        form, _, lawyer = split_data(data)

        client_email = form["clientEmail"]
        client_language = get_language(data)
        lawyer_name = lawyer.get("name", "")
        lawyer_id = lawyer.get("id", "")
        lang = "fr" if client_language.startswith("fr") else "en"

        template_path = os.path.join(TEMPLATES_DIR, lang, "Reply.html")

        with open(template_path, "r", encoding="utf-8") as f:
            html_body = f.read()

        html_body = html_body.replace("{{lawyerName}}", get_lawyer_string(lawyer_name, lawyer_id))

        outlook = COM.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = client_email
        mail.Subject = "Réponse - Allen Madelin" if lang == "fr" else "Reply - Allen Madelin"
        mail.HTMLBody = html_body
        mail.Display()
        
        focus_office_window(mail)
        print("Reply email draft opened in Outlook.")

    except Exception as e:
        alert_error(f"Error: {e}")
        raise

if __name__ == "__main__":
    draft_reply()

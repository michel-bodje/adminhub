from office_utils import *
from parse_json import *
import win32com.client as COM

def open_review_email():
    try:
        data = read_json("test.json")
        form, _, _ = split_data(data)

        client_email = form["clientEmail"]
        client_language = get_language(data)
        lang = "fr" if client_language.startswith("fr") else "en"

        template_path = os.path.join(TEMPLATES_DIR, lang, "Review.html")
        with open(template_path, "r", encoding="utf-8") as f:
            html_body = f.read()

        outlook = COM.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = client_email
        mail.Subject = "Commentaires Google - Allen Madelin" if lang == "fr" else "Google Review - Allen Madelin"
        mail.HTMLBody = html_body
        mail.Display()

        focus_office_window(mail)
        log("Review email draft opened in Outlook.")

    except Exception as e:
        alert_error(f"Error: {e}")
        raise

if __name__ == "__main__":
    open_review_email()

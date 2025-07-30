import win32com.client as COM
import os
from office_utils import *
from parse_json import *

def draft_contract():
    try:
        # Load JSON data
        data = read_json()
        form, _, lawyer = split_data(data)

        # Get PDF path from form data (passed from JavaScript)
        pdf_path = form.get("pdfPath", "")
        
        client_email = form.get("clientEmail", "")
        client_language = form.get("clientLanguage", "")
        deposit_amount = float(form.get("depositAmount", 0))
        lawyer_name = lawyer.get("name", "")
        lawyer_id = lawyer.get("id", "")
        
        # Determine language
        lang = "fr" if client_language == "Français" else "en"
        
        # Load email template
        template_path = os.path.join(TEMPLATES_DIR, lang, "Contract.html")
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template not found: {template_path}")
            
        with open(template_path, "r", encoding="utf-8") as f:
            html_body = f.read()
        
        # Calculate amounts and prepare content
        total_amount = add_taxes(deposit_amount, add_fof=True)
        lawyer_string = get_lawyer_string(lawyer_name, lawyer_id)
        
        # Replace placeholders in email template
        html_body = (html_body
                    .replace("{{depositAmount}}", f"{deposit_amount:.0f}")
                    .replace("{{totalAmount}}", f"{total_amount:.2f}")
                    .replace("{{lawyerName}}", lawyer_string))
        
        # Set email subject
        subject = ("Contrat de services - Allen Madelin" if lang == "fr" 
                  else "Contract of services - Allen Madelin")
        
        # Create Outlook email
        outlook_app = COM.Dispatch("Outlook.Application")
        mail = outlook_app.CreateItem(0)  # olMailItem
        
        mail.To = client_email
        mail.Subject = subject
        mail.HTMLBody = html_body
        
        # Attach PDF if path was provided and file exists
        if pdf_path and os.path.exists(pdf_path):
            mail.Attachments.Add(pdf_path)
            print(f"Attached PDF contract: {pdf_path}")
        elif pdf_path:
            print(f"PDF path provided but file not found: {pdf_path}")
        else:
            print("No PDF path provided, creating email without attachment")
        
        # Display the email draft
        mail.Display()
        focus_office_window(mail)
        
        print("Contract email created successfully")
        
    except Exception as e:
        error_msg = f"Error creating contract email: {str(e)}"
        print(error_msg)
        alert_error(error_msg)
        raise

if __name__ == "__main__":
    draft_contract()
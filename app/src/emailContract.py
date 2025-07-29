import win32com.client as COM
from office_utils import *
from parse_json import *

def draft_contract():
    try:
        # === Load JSON ===
        data = read_json()
        form, _, lawyer = split_data(data)

        # Check if PDF path was provided by frontend
        pdf_path = form.get("pdfPath", "")
        
        client_email = form.get("clientEmail", "")
        client_language = form.get("clientLanguage", "")
        deposit_amount = float(form.get("depositAmount", 0))
        lawyer_name = lawyer.get("name", "")
        lawyer_id = lawyer.get("id", "")
        
        # Determine language
        lang = "fr" if client_language == "Français" else "en"
        
        # Load template
        template_path = os.path.join(TEMPLATES_DIR, lang, "Contract.html")
        
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template not found: {template_path}")
            
        with open(template_path, "r", encoding="utf-8") as f:
            html_body = f.read()
        
        # Calculate amounts
        total_amount = add_taxes(deposit_amount, add_fof=True)
        lawyer_string = get_lawyer_string(lawyer_name, lawyer_id)
        
        # === Replace placeholders ===
        html_body = (html_body
                    .replace("{{depositAmount}}", f"{deposit_amount:.0f}")
                    .replace("{{totalAmount}}", f"{total_amount:.2f}")
                    .replace("{{lawyerName}}", lawyer_string))
        
        # Set subject
        subject = ("Contrat de services - Allen Madelin" if lang == "fr" 
                  else "Contract of services - Allen Madelin")
        
        # === Create email using Outlook ===
        outlook_app = COM.Dispatch("Outlook.Application")
        mail = outlook_app.CreateItem(0)  # olMailItem
        
        mail.To = client_email
        mail.Subject = subject
        mail.HTMLBody = html_body
        
        # Check for PDF attachment
        path_temp_file = os.path.join(ROOT_DIR, "data", "latest_contract_path.txt")
        
        if os.path.exists(path_temp_file):
            with open(path_temp_file, "r") as f:
                pdf_path = f.read().strip()
                
            if os.path.exists(pdf_path):
                mail.Attachments.Add(pdf_path)
                print(f"Attached PDF contract: {pdf_path}")
            else:
                print(f"PDF path recorded, but file no longer exists: {pdf_path}")
        else:
            print("No PDF was generated, skipping attachment.")
        
        # Display and focus
        mail.Display()  # Opens the draft
        focus_office_window(mail)  # Use existing focus function
        
        # === Clean up ===
        if os.path.exists(path_temp_file):
            os.remove(path_temp_file)
            
        print("Draft contract email created successfully.")
        
    except Exception as e:
        error_msg = f"Error: {str(e)}"
        print(error_msg)
        alert_error(error_msg)
        raise

if __name__ == "__main__":
    draft_contract()
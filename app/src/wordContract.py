import shutil
from datetime import datetime
from tkinter import Tk, filedialog
import pythoncom
import win32com.client as COM
from office_utils import *
from parse_json import *


def open_contract_draft():
    try:
        data = read_json()
        form, _, _ = split_data(data)

        client_name = form["clientName"]
        client_email = form["clientEmail"]
        client_language = get_language(data)
        contract_title = form.get("contractTitle", "")
        deposit_amount = float(form.get("depositAmount", 0))

        lang = "fr" if client_language.startswith("fr") else "en"
        template_path = os.path.join(TEMPLATES_DIR, lang, "Contract.docx")
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template not found: {template_path}")

        # Create temp file
        temp_doc_path = os.path.join(os.environ.get("TEMP", "/tmp"), f"Contract_{os.urandom(4).hex()}.docx")
        shutil.copy(template_path, temp_doc_path)

        # Compute replacements
        total_amount = add_taxes(deposit_amount, add_fof=True)
        formatted_deposit = f"{deposit_amount:.0f}"
        formatted_amount = f"{total_amount:.2f}"
        today = format_date(datetime.today(), lang)

        title_map = {
            "divorce": ("Représentation en divorce", "Representation in Divorce"),
            "estate": ("Représentation en droit des successions", "Representation in Estate Law"),
            "limited": ("Mandat Limité", "Limited Mandate")
        }
        if contract_title in title_map:
            title_text = title_map[contract_title][0 if lang == "fr" else 1]
        elif contract_title:
            title_text = contract_title  # raw, untranslated
        else:
            title_text = ""  # or raise an error if this should never happen

        replacements = {
            "{clientName}": client_name,
            "{contractTitle}": title_text,
            "{depositAmount}": formatted_deposit,
            "{totalAmount}": formatted_amount,
            "{date}": today
        }

        # Initialize Word COM
        pythoncom.CoInitialize()
        word = COM.DispatchEx("Word.Application")
        doc = word.Documents.Open(temp_doc_path)

        for placeholder, replacement in replacements.items():
            word_replace_text(doc, placeholder, replacement)
            
        word_hyperlink_email(doc, "{clientEmail}", client_email)

        # Show Word
        word.Visible = True
        focus_office_window(doc.ActiveWindow)

        # Save PDF dialog
        Tk().withdraw()
        default_filename = f"{'Contrat de services' if lang == 'fr' else 'Contract of services'}_{client_name.replace(' ', '-')}_{datetime.today().strftime('%Y-%m-%d')}.pdf"

        pdf_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save Contract as PDF",
            initialfile=default_filename
        )

        if pdf_path:
            doc.ExportAsFixedFormat(pdf_path, 17)  # 17 = wdExportFormatPDF
            print(f"Contract saved to PDF: {pdf_path}")
            doc.Close(False)
            word.Quit()
            os.remove(temp_doc_path)
            
            # Return path for frontend to capture
            return pdf_path
        else:
            call_cleaner_async(temp_doc_path)
    except Exception as e:
        alert_error(f"Error: {e}")
        raise
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass

if __name__ == "__main__":
    open_contract_draft()

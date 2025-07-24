import shutil
import pythoncom
import win32com.client as COM
from datetime import datetime
from tkinter import Tk, filedialog
from office_utils import *
from parse_json import *


def open_receipt():
    try:
        data = read_json()
        form, _, lawyer = split_data(data)

        client_name = form.get("clientName", "")
        client_language = get_language(data)
        payment_method = form.get("paymentMethod", "")
        deposit_amount = float(form.get("depositAmount", 0))
        receipt_reason = form.get("receiptReason", "")
        lawyer_id = lawyer.get("id", "")
        lawyer_name = lawyer.get("name", "")

        lang = "fr" if client_language.startswith("Français") else "en"
        template_path = os.path.join(TEMPLATES_DIR, lang, "Receipt.docx")
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Receipt template not found: {template_path}")

        # Create temp copy of template
        temp_doc_path = os.path.join(os.environ.get("TEMP", "/tmp"), f"Receipt_{os.urandom(4).hex()}.docx")
        shutil.copy(template_path, temp_doc_path)

        # Formating
        formatted_date = format_date(datetime.today(), lang)
        formatted_amount = f"{deposit_amount:.2f}"
        lawyer_string = get_lawyer_string(lawyer_name, lawyer_id)

        # Prepare reason text
        reason_text = {
            "consultation": "une consultation juridique" if lang == "fr" else "a legal consultation",
            "trust": "un paiement en fidéicommis" if lang == "fr" else "a trust payment"
        }.get(receipt_reason, receipt_reason)

        replacements = {
            "{user}": "Michel Assi-Bodje",
            "{reason}": reason_text,
            "{clientName}": client_name,
            "{paymentMethod}": payment_method,
            "{depositAmount}": formatted_amount,
            "{lawyerName}": lawyer_string,
            "{date}": formatted_date
        }
        
        # Init Word COM
        pythoncom.CoInitialize()
        word = COM.DispatchEx("Word.Application")
        doc = word.Documents.Open(temp_doc_path)

        for placeholder, replacement in replacements.items():
            word_replace_text(doc, placeholder, replacement)

        # Show Word window and focus
        word.Visible = True
        focus_office_window(doc.ActiveWindow)

        # Print dialog (File > Print)
        word.Dialogs(88).Show()  # 88 = wdDialogFilePrint

        # Save as PDF dialog
        Tk().withdraw()
        default_filename = f"{datetime.today():%Y-%m-%d}_{client_name.replace(' ', '-')}.pdf"
        pdf_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save Receipt as PDF",
            initialfile=default_filename
        )

        if pdf_path:
            doc.ExportAsFixedFormat(pdf_path, 17)  # 17 = wdExportFormatPDF
            print("PDF saved to:", pdf_path)
            doc.Close(False)
            word.Quit()
            os.remove(temp_doc_path)
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
    open_receipt()

from office_utils import *
from parse_json import *
import shutil
import tempfile
import pythoncom
from datetime import datetime
from tkinter import Tk, filedialog

def process_word_receipt(data):
    try:
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

        # Create temp file in a reliable location - FIXED
        temp_dir = tempfile.gettempdir()
        temp_doc_path = os.path.join(temp_dir, f"Receipt_{os.urandom(4).hex()}.docx")
        
        # Ensure temp directory exists
        os.makedirs(temp_dir, exist_ok=True)
        shutil.copy(template_path, temp_doc_path)
        
        # Verify the temp file was created
        if not os.path.exists(temp_doc_path):
            raise FileNotFoundError(f"Failed to create temporary file: {temp_doc_path}")

        # Formatting
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
        
        # Init Word COM with better error handling
        pythoncom.CoInitialize()
        try:
            word = COM.DispatchEx("Word.Application")
            word.Visible = False  # Keep hidden initially
            doc = word.Documents.Open(temp_doc_path)
        except Exception as word_error:
            raise Exception(f"Failed to open Word document: {str(word_error)}")

        # Process replacements
        try:
            for placeholder, replacement in replacements.items():
                word_replace_text(doc, placeholder, replacement)
        except Exception as replace_error:
            raise Exception(f"Failed to process document replacements: {str(replace_error)}")

        # Show Word window after processing
        word.Visible = True
        focus_office_window(doc.ActiveWindow)

        # Print dialog (File > Print)
        try:
            word.Dialogs(88).Show()  # 88 = wdDialogFilePrint
        except Exception as print_error:
            print(f"Warning: Print dialog failed: {print_error}", file=sys.stderr)

        # Save as PDF dialog with better error handling
        try:
            root = Tk()
            root.withdraw()
            root.lift()  # Bring to front
            root.attributes('-topmost', True)  # Keep on top
            
            default_filename = f"{datetime.today():%Y-%m-%d}_{client_name.replace(' ', '-')}.pdf"
            
            pdf_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                title="Save Receipt as PDF",
                initialfile=default_filename
            )
            root.destroy()
        except Exception as dialog_error:
            raise Exception(f"Failed to show save dialog: {str(dialog_error)}")

        if pdf_path:
            try:
                # Validate and sanitize the PDF path (similar to contract)
                pdf_path = os.path.normpath(pdf_path)
                pdf_dir = os.path.dirname(pdf_path)
                
                # Check if directory exists, create if it doesn't
                if not os.path.exists(pdf_dir):
                    os.makedirs(pdf_dir, exist_ok=True)
                
                # Check if we can write to the directory
                if not os.access(pdf_dir, os.W_OK):
                    raise Exception(f"No write permission to directory: {pdf_dir}")
                
                # Sanitize filename
                filename = os.path.basename(pdf_path)
                invalid_chars = '<>:"/\\|?*'
                for char in invalid_chars:
                    if char in filename:
                        filename = filename.replace(char, '_')
                
                pdf_path = os.path.join(pdf_dir, filename)
                
                print(f"Attempting to export PDF to: {pdf_path}", file=sys.stderr)
                
                doc.ExportAsFixedFormat(pdf_path, 17)  # 17 = wdExportFormatPDF
                              
            except Exception as export_error:
                raise Exception(f"PDF export failed: {str(export_error)}")
            finally:
                # Always clean up Word
                try:
                    doc.Close(False)
                    word.Quit()
                
                    return {
                        "status": "success",
                        "message": "Word receipt successfully created.",
                        "pdf_path": pdf_path
                    }
                except:
                    pass
                # Clean up temp file
                try:
                    if os.path.exists(temp_doc_path):
                        os.remove(temp_doc_path)
                except:
                    # If immediate cleanup fails, use the cleaner
                    call_cleaner_async(temp_doc_path)
        else:
            # User cancelled - clean up
            try:
                doc.Close(False)
                word.Quit()
            except:
                pass
            call_cleaner_async(temp_doc_path)

            return {
                "status": "cancelled",
                "message": "User cancelled receipt export to PDF."
            }
            
    except Exception as e:
        alert_error(f"Error: {e}")
        return {
            "status": "error",
            "error": str(e),
            "message": f"Failed to create receipt: {str(e)}"
        }
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass

# Backward compatibility
def main():
    try:
        data = read_json()
        result = process_word_receipt(data)
        if result.get("error"):
            print(f"Error: {result['error']}")
        else:
            print(result.get("message", "Word receipt processed"))
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
import shutil
import json
import sys
import os
import tempfile
from datetime import datetime
from tkinter import Tk, filedialog
import pythoncom
import win32com.client as COM
from office_utils import *
from parse_json import *

def process_word_contract(data):
    try:
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

        # Create temp file in a reliable location
        temp_dir = tempfile.gettempdir()
        temp_doc_path = os.path.join(temp_dir, f"Contract_{os.urandom(4).hex()}.docx")
        
        # Ensure temp directory exists
        os.makedirs(temp_dir, exist_ok=True)
        shutil.copy(template_path, temp_doc_path)
        
        # Verify the temp file was created
        if not os.path.exists(temp_doc_path):
            raise FileNotFoundError(f"Failed to create temporary file: {temp_doc_path}")

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
            title_text = contract_title
        else:
            title_text = ""

        replacements = {
            "{clientName}": client_name,
            "{contractTitle}": title_text,
            "{depositAmount}": formatted_deposit,
            "{totalAmount}": formatted_amount,
            "{date}": today
        }

        # Initialize Word COM with better error handling
        pythoncom.CoInitialize()
        try:
            word = COM.DispatchEx("Word.Application")
            word.Visible = False  # Keep hidden initially to avoid display issues
            doc = word.Documents.Open(temp_doc_path)
        except Exception as word_error:
            raise Exception(f"Failed to open Word document: {str(word_error)}")

        # Process replacements
        try:
            for placeholder, replacement in replacements.items():
                word_replace_text(doc, placeholder, replacement)
                
            word_hyperlink_email(doc, "{clientEmail}", client_email)
        except Exception as replace_error:
            raise Exception(f"Failed to process document replacements: {str(replace_error)}")

        # Show Word after processing
        word.Visible = True
        focus_office_window(doc.ActiveWindow)

        # Save PDF dialog with better error handling
        try:
            root = Tk()
            root.withdraw()
            root.lift()  # Bring to front
            root.attributes('-topmost', True)  # Keep on top
            
            default_filename = f"{'Contrat de services' if lang == 'fr' else 'Contract of services'}_{client_name.replace(' ', '-')}_{datetime.today().strftime('%Y-%m-%d')}.pdf"

            pdf_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                title="Save Contract as PDF",
                initialfile=default_filename
            )
            root.destroy()
        except Exception as dialog_error:
            raise Exception(f"Failed to show save dialog: {str(dialog_error)}")

        if pdf_path:
            try:
                # Validate and sanitize the PDF path
                pdf_path = os.path.normpath(pdf_path)  # Normalize path separators
                pdf_dir = os.path.dirname(pdf_path)
                
                # Check if directory exists, create if it doesn't
                if not os.path.exists(pdf_dir):
                    os.makedirs(pdf_dir, exist_ok=True)
                
                # Check if we can write to the directory
                if not os.access(pdf_dir, os.W_OK):
                    raise Exception(f"No write permission to directory: {pdf_dir}")
                
                # Ensure the filename doesn't have invalid characters
                filename = os.path.basename(pdf_path)
                invalid_chars = '<>:"/\\|?*'
                for char in invalid_chars:
                    if char in filename:
                        filename = filename.replace(char, '_')
                
                # Reconstruct the full path with sanitized filename
                pdf_path = os.path.join(pdf_dir, filename)
                
                print(f"Attempting to export PDF to: {pdf_path}", file=sys.stderr)
                
                doc.ExportAsFixedFormat(pdf_path, 17)  # 17 = wdExportFormatPDF
                
                # Always clean up Word
                doc.Close(False)
                word.Quit()

                return {
                    "status": "success",
                    "message": "Word contract successfully created.",
                    "pdf_path": pdf_path
                }

            except Exception as export_error:
                raise Exception(f"PDF export failed: {str(export_error)}")
            finally:
                # Clean up temp file
                try:
                    if os.path.exists(temp_doc_path):
                        os.remove(temp_doc_path)
                except:
                    pass
        else:
            # User cancelled - clean up
            call_cleaner_async(temp_doc_path)

            return {
                "status": "Cancelled",
                "message": "User cancelled contract export to PDF.",
            }
            
    except Exception as e:
        print(json.dumps({"error": str(e)}), file=sys.stderr)
        alert_error(f"Error: {e}")
        raise
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass

# Backward compatibility
def main():
    try:
        data = read_json()
        result = process_word_contract(data)
        if result.get("error"):
            print(f"Error: {result['error']}")
        else:
            print(result.get("message", "Word contract processed"))
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
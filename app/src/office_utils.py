from config import *
import re
import locale
from time import sleep
import win32clipboard as clipboard
from win32com import client as COM
from pywinauto.application import Application
from pywinauto.findwindows import find_windows

# ----------------------------
# Various Utilities
# ----------------------------

def add_taxes(amount, add_fof=False):
    if not isinstance(amount, (int, float)):
        raise ValueError("Amount must be a number.")
    total = amount * (1 + 0.05 + 0.09975)
    if add_fof:
        total += 100
    return total

def get_ordinal_suffix(n):
    if 11 <= n % 100 <= 13:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")

def format_date(slot, lang):
    """ Format date based on language. """
    if lang == "en":
        day = slot.day
        suffix = get_ordinal_suffix(day)
        return slot.strftime(f"%A %B {day}{suffix}, %Y")
    else:
        # French long format (e.g., mardi 20 juillet 2025)
        try:
            locale.setlocale(locale.LC_TIME, 'fr_FR')
        except locale.Error:
            print("Could not set locale to French. Date formatting may be incorrect.")
        date_string = slot.strftime("%A %d %B %Y")
        locale.setlocale(locale.LC_TIME, '') # Reset to default locale
        return date_string

def format_time(slot, lang):
    """ Format time based on language: 12-hour for English, 24-hour for French. """
    return slot.strftime("%I:%M %p" if lang == "en" else "%H:%M").lstrip("0").replace(" 0", " ")

def get_lawyer_string(name, lawyer_id) -> str:
    """ Format lawyer name with prefix if applicable. """
    if lawyer_id not in ("AR", "MG", "PM"):
        return f"Me {name}"
    return name


# ----------------------------
# Window Focus
# ----------------------------

def focus_office_window(item):
    """Brings the created Office window to the foreground."""
    try:
        inspector = item.GetInspector
        inspector.Activate()  # This opens and focuses the window
        return True
    except Exception as e:
        print(f"Failed to activate Inspector: {e}")
        return False


# ----------------------------
# Teams Meeting Block
# ----------------------------

def connect_to_meeting_window():
    """Connect to an Outlook Meeting/Appointment window by title."""
    hwnds = find_windows(title_re=r".*(- Meeting|- Appointment).*")
    if not hwnds:
        raise RuntimeError("No Meeting window found")
    app = Application(backend="uia").connect(handle=hwnds[0])
    win = app.window(handle=hwnds[0])
    win.wait("visible enabled ready", timeout=10)
    return win

def click_teams_meeting_button():
    """Find and click the Teams Meeting button on the ribbon."""
    meeting_win = connect_to_meeting_window()
    # Try to locate the button by visible label
    btn = meeting_win.child_window(
        title_re=r"(Teams Meeting|Réunion Teams)", 
        control_type="Button"
    )
    btn.wait("visible enabled ready", timeout=5)
    btn.invoke()  # same as pressing it
    print("[OK] Clicked Teams Meeting button")

def copy_teams_block(doc):
    try:
        find = doc.Content.Find
        find.Text = "Microsoft Teams"
        if not find.Execute():
            print("[WARN] Teams block not found")
            return None

        found_para = find.Parent.Paragraphs(1)

        # Find index by matching paragraph start position
        start_para_index = None
        for i in range(1, doc.Paragraphs.Count + 1):
            para = doc.Paragraphs(i)
            if para.Range.Start == found_para.Range.Start:
                start_para_index = i
                break

        if start_para_index is None:
            print("[ERROR] Could not find paragraph index")
            return None

        total_paras = doc.Paragraphs.Count
        end_para_index = min(start_para_index + 3, total_paras)

        start_para_range = doc.Paragraphs(start_para_index).Range
        end_para_range = doc.Paragraphs(end_para_index).Range

        block_rng = doc.Range(Start=start_para_range.Start, End=end_para_range.End)
        block_rng.Copy()
        sleep(0.5)
        print("[OK] Teams block copied to clipboard")
        
    except Exception as e:
        print(f"[ERROR] Failed to copy Teams block: {e}")
        return None

def get_teams_block():
    try:
        # Read clipboard HTML format
        html_bytes = get_clipboard_html()
        if html_bytes:
            html_str = html_bytes.decode('utf-8')
            html_str = extract_html_fragment(html_str)
            html_str = clean_html_fragment(html_str)
            html_str = convert_p_to_br(html_str)

            print("[OK] Teams block copied and HTML extracted from clipboard")
            return html_str
        else:
            print("[WARN] Failed to get HTML from clipboard after copying Teams block")
            return None
    except Exception as e:
        print(f"[ERROR] Failed to process Teams block: {e}")
        return None

def get_clipboard_html():
    clipboard.OpenClipboard()
    try:
        CF_HTML = clipboard.RegisterClipboardFormat("HTML Format")
        if clipboard.IsClipboardFormatAvailable(CF_HTML):
            html = clipboard.GetClipboardData(CF_HTML)
            return html
        return None
    finally:
        clipboard.CloseClipboard()

def extract_html_fragment(full_html: str) -> str:
    start_marker = "<!--StartFragment-->"
    end_marker = "<!--EndFragment-->"
    
    start_index = full_html.find(start_marker)
    if start_index == -1:
        return None  # start fragment not found
    
    start_index += len(start_marker)  # move past the marker
    end_index = full_html.find(end_marker, start_index)
    if end_index == -1:
        return None  # end fragment not found
    
    fragment = full_html[start_index:end_index]
    return fragment.strip()

def clean_html_fragment(fragment: str) -> str:
    # Split into lines
    lines = fragment.splitlines()
    # Remove empty or whitespace-only lines
    lines = [line for line in lines if line.strip() != '']
    cleaned = '\n'.join(lines)

    # Add a newline before every opening <p> tag (if not already at start of line)
    cleaned = re.sub(r'(?<!\n)(<p[^>]*>)', r'\n\1', cleaned)

    # Strip leading/trailing whitespace from whole string
    cleaned = cleaned.strip()
    return cleaned

def convert_p_to_br(fragment: str) -> str:
    # Remove closing </p> tags, replace opening <p> with nothing or just add <br>
    fragment = fragment.replace('</p>', '<br>')
    # Remove opening <p> tags (optionally with attributes)
    fragment = re.sub(r'<p[^>]*>', '', fragment)
    return fragment.strip()


# ----------------------------
# Word Automation Helpers
# ----------------------------

def word_replace_text(doc, placeholder, replacement):
    """Replace all occurrences of a placeholder in the Word document using Word COM."""
    find = doc.Content.Find

    find.ClearFormatting()
    find.Replacement.ClearFormatting()

    find.Text = placeholder
    find.Replacement.Text = replacement

    find.Forward = True
    find.Wrap = 1  # wdFindContinue
    find.Format = False
    find.MatchCase = False
    find.MatchWholeWord = False
    find.MatchWildcards = False
    find.MatchSoundsLike = False
    find.MatchAllWordForms = False

    # THIS is the crucial part: explicit parameters
    find.Execute(
        FindText=placeholder,
        MatchCase=False,
        MatchWholeWord=False,
        MatchWildcards=False,
        MatchSoundsLike=False,
        MatchAllWordForms=False,
        Forward=True,
        Wrap=1,
        Format=False,
        ReplaceWith=replacement,
        Replace=2  # wdReplaceAll
    )

def word_hyperlink_email(doc, placeholder: str, email: str):
    """Replace email placeholder with a hyperlink to an email address."""
    range_obj = doc.Content
    find = range_obj.Find

    find.ClearFormatting()
    find.Text = placeholder

    if find.Execute():
        range_obj.Hyperlinks.Add(
            Anchor=range_obj,
            Address=f"mailto:{email}",
            SubAddress="",
            ScreenTip="Email Client",
            TextToDisplay=email
        )
    else:
        log(f"[Error] Could not find placeholder for email: {placeholder}")


# ----------------------------
# Temp Cleanup
# ----------------------------

def call_cleaner_async(temp_doc_path: str):
    """
    Launches cleaner process in background without blocking.
    
    Now uses direct module import instead of subprocess to work
    properly with PyInstaller bundled executables.
    """
    try:
        # Import the cleanup function directly as a module
        from cleanTempDoc import cleanup_temp_doc_async
        
        # Call it directly - no subprocess needed!
        cleanup_temp_doc_async(temp_doc_path, timeout_minutes=30)
        
        log(f"Started background cleanup for temp document: {temp_doc_path}")
        
    except ImportError as e:
        log(f"Could not import cleanup module: {e}")
        alert_error("Cleanup functionality unavailable - temp files may need manual deletion")
    except Exception as e:
        log(f"Error starting cleanup: {e}")
        alert_error(f"Could not start file cleanup: {e}")
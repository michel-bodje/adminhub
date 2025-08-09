from config import *
import re
import locale
import win32com.client as COM
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

def get_teams_meeting_block():
    """Return the 4-line Teams Meeting block from the appointment body."""
    outlook = COM.Dispatch("Outlook.Application")
    inspector = outlook.ActiveInspector()
    if not inspector:
        raise RuntimeError("No active inspector found")
    appt = inspector.CurrentItem
    body = appt.Body
    # Regex to match the 4 lines
    pattern = re.compile(
        r"(Microsoft Teams.*?\n"                                     # First line
        r"(?:Join the meeting now|Participer à la réunion maintenant).*?\n"  # Second line
        r"(?:Meeting ID|ID de réunion)\s*:\s*.*?\n"                  # Third line
        r"(?:Passcode|Code d’accès)\s*:\s*.*?)(?:\n|$)",             # Fourth line
        re.IGNORECASE | re.DOTALL
    )
    match = pattern.search(body)
    if match:
        return match.group(1).strip()
    else:
        print("[WARN] Teams meeting block not found")
        return None


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
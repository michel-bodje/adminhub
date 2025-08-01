import locale
from config import *

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
            locale.setlocale(locale.LC_TIME, 'fr_FR.UTF-8')
        except locale.Error:
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
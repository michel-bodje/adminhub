import os
import sys
import ctypes

def log(message):
    print(f"[AdminHub] {message}")

"""
# --------------------------------------------
# MessageBoxW Flag Reference
# --------------------------------------------
# Icons:
#   0x10 - Stop (❌)
#   0x20 - Question (❓)
#   0x30 - Warning (⚠️)
#   0x40 - Information (ℹ️)

# Buttons:
#   0x00 - OK
#   0x01 - OK / Cancel
#   0x03 - Yes / No / Cancel
#   0x04 - Yes / No
#   0x05 - Retry / Cancel

# Return Values:
#   1 - OK
#   2 - Cancel
#   3 - Abort
#   4 - Retry
#   5 - Ignore
#   6 - Yes
#   7 - No

# Examples:
#   0x21 = Question icon + OK/Cancel
#   0x30 = Warning icon + OK
#   0x34 = Warning icon + Yes/No
#   0x40 = Info icon + OK
#   0x53 = Stop icon + Yes/No/Cancel
# --------------------------------------------
"""
# Function to display a message box
# This function can be used to show different types of message boxes
def message_box(message, title="AdminHub", flags=0x40):
    log(message)
    return ctypes.windll.user32.MessageBoxW(0, message, title, flags | 0x1000)  # 0x1000 = MB_TOPMOST, ensures the message box is always on top

def alert_info(msg, title="Success - AdminHub"):
    message_box(msg, title, 0x40)

def alert_warning(msg, title="Warning - AdminHub"):
    message_box(msg, title, 0x30)

def alert_error(msg, title="Error - AdminHub"):
    message_box(msg, title, 0x10)

def confirm_continue(msg, title="Continue? - AdminHub"):
    result = message_box(msg, title, 0x24)  # Question icon + Yes/No
    return result == 6  # 6 = Yes

def get_root_path():
    """ Returns the root path of the application.
    This function checks if the application is running as a bundled executable
    (e.g., created with PyInstaller) or as a regular Python script.
    If bundled, it returns the temporary folder where the bundled files are located.
    If running as a script, it returns the absolute path to the project root directory.
    """
    if getattr(sys, 'frozen', False):
        # temp folder with bundled files
        return getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    else:
        # else get to the project root
        return os.path.abspath(os.path.join(os.path.dirname(__file__), '../../'))

ROOT_DIR = get_root_path()
SRC_DIR = os.path.join(ROOT_DIR, 'app', 'src')
WEB_DIR = os.path.join(ROOT_DIR, 'app', 'web')
TEMPLATES_DIR = os.path.join(ROOT_DIR, 'app', 'templates')
INDEX_HTML = os.path.join(WEB_DIR, 'index.html')
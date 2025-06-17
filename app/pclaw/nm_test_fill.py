from pywinauto.application import Application
from pywinauto.findwindows import find_windows
from pywinauto.timings import wait_until_passes
from time import sleep

# Connect to open PCLaw instance
hwnds = find_windows(title_re=".*PCLaw.*")
app = Application(backend="uia").connect(handle=hwnds[0])
main_win = app.window(handle=hwnds[0])
main_win.set_focus()

# Open "New Matter" if not already open
main_win.child_window(title="New Matter", control_type="Button").click_input()
sleep(1)  # wait for dialog to open

# Locate the New Matter dialog
new_matter_dlg = app.window(title_re=".*New Matter.*")
wait_until_passes(5, 1, lambda: new_matter_dlg.exists(timeout=1) and new_matter_dlg.is_visible())

# Set focus
new_matter_dlg.set_focus()

# Fill in fields using auto_id values
fields = {
    "259335784": "Mr.",          # Title
    "259336688": "Michel",       # First
    "259358384": "Assi",         # Middle
    "259337592": "Bodje",        # Last
    "259345728": "DH",           # Responsible Lawyer
    # "259354768": "2025-001",     # Matter
    # "259352960": "Client123",    # Client
    "259357480": "Civil",        # Type of Law
    "259333976": "125",          # Default Rate
    "259334880": "Case of Test" # Description
}

for auto_id, value in fields.items():
    try:
        field = new_matter_dlg.child_window(auto_id=auto_id, control_type="Edit")
        field.set_edit_text(value)
        sleep(0.1)  # slight pause for safety
    except Exception as e:
        print(f"Could not fill field {auto_id}: {e}")

# Optional: press OK
"""
try:
    new_matter_dlg.child_window(title="OK", control_type="Button").click_input()
except Exception as e:
    print(f"Could not click OK: {e}")
"""
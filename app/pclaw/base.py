from pywinauto.application import Application
from pywinauto.findwindows import find_windows
from pywinauto.timings import wait_until_passes
from time import sleep

# Connect to open PCLaw instance
hwnds = find_windows(title_re=".*PCLaw.*")
app = Application(backend="uia").connect(handle=hwnds[0])
main_win = app.window(handle=hwnds[0])
main_win.set_focus()

# Click a toolbar button by its visible title
def click_toolbar_button(title):
    print(f"Trying to click toolbar button: {title}")
    try:
        btn = main_win.child_window(title=title, control_type="Button")
        wait_until_passes(5, 1, lambda: btn.exists(timeout=1) and btn.is_enabled())
        btn.click_input()
        print(f"Clicked: {title}")
    except Exception as e:
        print(f"Error clicking '{title}':", e)

"""
Example usage of toolbar button clicks.
Each button will be clicked with a 2 second pause in between.

click_toolbar_button("New Matter")
sleep(2)

click_toolbar_button("Contact Manager")
sleep(2)

click_toolbar_button("Conflict Search")
"""

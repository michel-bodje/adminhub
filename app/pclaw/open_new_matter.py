from pywinauto.application import Application
from pywinauto.findwindows import find_windows
import time

# Connect to PCLaw
hwnds = find_windows(title_re=".*PCLaw.*")
app = Application(backend="uia").connect(handle=hwnds[0])
main_win = app.window(handle=hwnds[0])
main_win.set_focus()

# Click "New Matter"
main_win.child_window(title="New Matter", control_type="Button").click_input()

# Wait and take action in the new dialog (if needed)
time.sleep(1)
new_matter_dlg = app.window(title_re=".*New Matter.*")
new_matter_dlg.set_focus()

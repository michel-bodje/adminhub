from pywinauto.application import Application
from pywinauto.findwindows import find_windows
from pywinauto.timings import wait_until_passes
import datetime

# Optional: give each run a timestamped log
timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
output_file = f"pclaw_ui_snapshot_{timestamp}.txt"

# Connect to the running PCLaw instance (update if needed)
try:
    hwnds = find_windows(title_re=".*PCLaw.*")
    app = Application(backend="uia").connect(handle=hwnds[0])
except Exception as e:
    print("Could not find PCLaw window:", e)
    exit()

# Get top window
main_win = app.window(handle=hwnds[0])
main_win.set_focus()

# Wait for window to be ready (adjust timeout if needed)
wait_until_passes(5, 1, lambda: main_win.exists(timeout=1) and main_win.is_visible())

# Print a summary to file
with open(output_file, "w", encoding="utf-8") as f:
    f.write(f"--- UI Snapshot: {main_win.window_text()} ---\n")
    f.write(f"Timestamp: {timestamp}\n")
    f.write(f"{'-'*40}\n\n")
    main_win.print_control_identifiers(filename=f.name)

print(f"Snapshot saved to {output_file}")

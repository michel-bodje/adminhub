from pywinauto import Application
from pywinauto.findwindows import find_windows
from base import *
import sys

if __name__ == "__main__":
    dlg = open_dialog(connect_to_pclaw(), "New Matter", "New Matter")
    go_to_billing(dlg)
    click_billing_checkbox(dlg)
    
    with open("hierarchy.txt", "w", encoding="utf-8") as f:
        dlg.print_control_identifiers(filename=f.name)
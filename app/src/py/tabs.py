from pywinauto import Application
from pywinauto.findwindows import find_windows
from base import *

if __name__ == "__main__":
    dlg = open_dialog(connect_to_pclaw(), "New Matter", "New Matter")
    go_to_custom(dlg)
    
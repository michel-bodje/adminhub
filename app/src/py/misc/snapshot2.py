from base import *
import sys

win = connect_to_pclaw()

new_matter_dialog()

dlg = get_dialog(win, "New Matter")

# setup a ui dialog window
go_to_billing(dlg)

# print ui hierarchy
with open("hierarchy.txt", "w", encoding="utf-8") as f:
    dlg.print_control_identifiers(filename=f.name)

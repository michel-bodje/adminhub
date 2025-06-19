from base import *
from parse_json import *

# Define the specific fields for New Matter
fields = load_consultation_fields()
lang = get_language()

dlg = open_dialog(connect_to_pclaw(), "New Matter", "New Matter")
label_input_map = build_label_input_map(dlg, fields)

# Fill main tab
fill_form_fields(label_input_map, fields)

# Only fill billing tab if language is French
if lang.startswith("fr"):
    go_to_billing(dlg)
    fill_billing_tab(dlg)

# Fill all n/a in Custom tab
go_to_custom(dlg)
fill_custom_tab(dlg)

# Optionally press OK
# dlg.child_window(title="OK", control_type="Button").click_input()

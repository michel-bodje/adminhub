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
    billing_pane = get_tab_pane(dlg, "Billing")
    if billing_pane:
        fill_billing_tab(billing_pane)

# Fill all n/a in Custom tab
custom_pane = get_tab_pane(dlg, "Custom")
if custom_pane:
    fill_custom_tab_fields(custom_pane)

# Optionally press OK
# dlg.child_window(title="OK", control_type="Button").click_input()

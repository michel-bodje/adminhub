from base import connect_to_pclaw, open_dialog, build_label_input_map, fill_form_fields
from parse_json import load_consultation_fields

# Define the specific fields for New Matter
fields = load_consultation_fields()

main_win = connect_to_pclaw()
new_matter_dlg = open_dialog(main_win, "New Matter", "New Matter")
label_input_map = build_label_input_map(new_matter_dlg, fields)
fill_form_fields(label_input_map, fields)

# Optionally press OK
# new_matter_dlg.child_window(title="OK", control_type="Button").click_input()

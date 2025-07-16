from base import *
from parse_json import *

def main():
    win = connect_to_pclaw()
    win.set_focus()

    new_matter_dialog()
    dlg = get_dialog(win, "New Matter")

    data = read_stdin_json()
  
    # Define the specific fields for New Matter
    fields = load_consultation_fields(data)
    label_input_map = build_label_input_map(dlg, fields)

    # Fill main tab
    fill_form_fields(label_input_map, fields)

    # Only fill billing tab if language is French
    if get_language(data).startswith("fr"):
        go_to_billing(dlg)
        fill_billing_tab_by_override(dlg)

    # Fill all n/a in Custom tab
    go_to_custom(dlg)
    fill_custom_tab(dlg)

    # Back to main tab
    go_to_main(dlg)

    # Optionally press OK
    # dlg.child_window(title="OK", control_type="Button").click_input()

if __name__ == "__main__":
    main()
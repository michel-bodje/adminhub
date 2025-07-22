from pclaw import *
from parse_json import *

def main():
    app = connect_to_pclaw()
    app.set_focus()

    new_matter_dialog()
    dlg = get_dialog(app, "New Matter")

    data = read_stdin_json()
  
    # Define the specific fields for New Matter
    fields = load_consultation_fields(data)
    
    fill_main_tab(fields)

    # Only fill billing tab if language is French
    if get_language(data).startswith("fr"):
        go_to_billing(dlg)
        fill_billing_tab(dlg)

    # Fill all n/a in Custom tab
    go_to_custom(dlg)
    fill_custom_tab(dlg)

    # Back to main tab
    go_to_main(dlg)

    # Optionally press OK
    # dlg.child_window(title="OK", control_type="Button").click_input()

if __name__ == "__main__":
    main()
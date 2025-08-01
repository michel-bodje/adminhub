from pclaw import *
from parse_json import *

def process_new_matter(data):
    """ Opens the New Matter dialog in PCLaw. """
    try:
        app = connect_to_pclaw()
        app.set_focus()
        sleep(1)

        new_matter_dialog()

        dlg = get_dialog(app, "New Matter")

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

        # Final confirmation
        sleep(1)

        alert_info("New matter created successfully in PCLaw.")

        return {
            "status": "success",
            "message": f"Successfully opened new matter",
        }
        
    except Exception as e:
        log(f"Error opening matter: {str(e)}")
        return {"error": str(e)}

def main():
    """Backward compatibility for standalone execution"""
    try:
        data = read_json()
        result = process_new_matter(data)
        if result.get("error"):
            print(f"Error: {result['error']}")
        else:
            print(result.get("message", "Matter opened successfully"))
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
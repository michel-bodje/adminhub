from parse_timesheets import DH_parse_timesheet, DH_record_time_entry, safe_correct
from pclaw import *
from parse_json import read_json

def startup():
    """ Connects to PCLaw and sets focus. """
    app = connect_to_pclaw()
    app.set_focus()
    sleep(3)  # Allow time for PCLaw to set focus
    return app

def enter_time():
    """ Main function to parse and record time entries from the time sheet. """
    data = read_json()
    
    form = data["form"]
    path = form.get("filePath", "")
    
    if not path:
        alert_error("File path not provided in the JSON data.")
        print("File path not provided in the JSON data.")
        return

    entries = DH_parse_timesheet(path)

    if not entries:
        alert_error("No valid time entries found in the sheet.")
        print("No valid time entries found in the sheet.")
        return

    continue_loop = True
    for entry in entries:
        if not continue_loop:
            break

        if not entry.recorded:
            saved = DH_record_time_entry(entry, path)
            if not saved:
                alert_info("User cancelled time entry. Exiting loop.")
                continue_loop = False

    # For testing, let's just record the first entry
    # entry = entries[0]
    #correction = safe_correct(entry.description)
    #if correction:
    #    print(f"Corrected description: {correction}")
    # DH_record_time_entry(entry, path)

def main():
    startup()
    enter_time()

if __name__ == "__main__":
    main()
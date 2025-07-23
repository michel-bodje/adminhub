from parse_timesheets import DH_parse_timesheet, DH_record_time_entry
from pclaw import *

def startup():
    """ Connects to PCLaw and sets focus. """
    app = connect_to_pclaw()
    app.set_focus()
    return app

def DH_enter_time():
    """ Main function to parse and record time entries from the time sheet. """
    # Path to the time sheet file
    # TODO: file picker
    path = r"\\AMNAS\amlex\Admin\Dorin Holban\DHO Time\DHO Time 2025\DHO Time - 2025-07.xlsx"

    entries = DH_parse_timesheet(path)
    if not entries:
        alert_error("No valid time entries found in the sheet.")
        print("No valid time entries found in the sheet.")
        return

    """
    for entry in entries:
        if not entry.recorded:
            record_time_entry(entry, path)
            print(f"Recorded: {entry.client} - {entry.description} on {entry.date} for {entry.time_spent} hours")
    """
    # For testing, let's just record the first entry
    entry = entries[0]
    DH_record_time_entry(entry, path)

def main():
    startup()
    DH_enter_time()

if __name__ == "__main__":
    main()
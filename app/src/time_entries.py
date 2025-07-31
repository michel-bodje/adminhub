from parse_timesheets import DH_parse_timesheet, DH_record_time_entry, safe_correct
from pclaw import *

def enter_time(data):
    """ Main function to parse and record time entries from the time sheet. """
    path = data.get("filePath", "")
    if not path:
        alert_error("File path not provided in the JSON data.")
        print("File path not provided in the JSON data.")
        return
    
    lawyer_id = data.get("lawyerId", "")
    if not lawyer_id:
        alert_error("Lawyer ID not provided in the JSON data.")
        print("Lawyer ID not provided in the JSON data.")
        return
    
    if lawyer_id == "DH":
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

    if lawyer_id == "TG":
        alert_error("TG time entries are not implemented yet.")
        pass

def main():
    _, data = startup()
    enter_time(data)

if __name__ == "__main__":
    main()
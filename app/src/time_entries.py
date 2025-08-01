from parse_timesheets import DH_parse_timesheet, DH_record_time_entry, safe_correct
from pclaw import *

def process_time_entries(data):
    """ Main function to parse and record time entries from the time sheet. """
    try:
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

        return {
            "status": "success",
            "message": f"Successfully recorded time entries",
            "lawyer_ID": lawyer_id,
            "file_path": path
        }
        
    except Exception as e:
        log(f"Error recording time entries: {str(e)}")
        return {"error": str(e)}    

def main():
    """Backward compatibility for standalone execution"""
    try:
        data = read_json()
        result = process_time_entries(data)
        if result.get("error"):
            print(f"Error: {result['error']}")
        else:
            print(result.get("message", "Time recorded successfully"))
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
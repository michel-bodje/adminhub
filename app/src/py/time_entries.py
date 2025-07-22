from parse_time_sheet import parse_DH_time_sheet, record_time_entry
from pclaw import *

def main():
    path = r"\\AMNAS\amlex\Admin\Dorin Holban\DHO Time\DHO Time 2025\DHO Time - 2025-07.xlsx"
    entries = parse_DH_time_sheet(path)
    
    app = connect_to_pclaw()
    app.set_focus()

    """
    for entry in entries:
        if not entry.recorded:
            record_time_entry(entry, path)
            print(f"Recorded: {entry.client} - {entry.description} on {entry.date} for {entry.time_spent} hours")
    """

    # For testing, let's just record the first entry
    entry = entries[0]
    
    record_time_entry(entry, path)
    print(f"Recorded: {entry.client} - {entry.description} on {entry.date} for {entry.time_spent} hours")

if __name__ == "__main__":
    main()
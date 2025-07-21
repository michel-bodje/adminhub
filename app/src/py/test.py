from parse_time_sheet import parse_DH_time_sheet

if __name__ == "__main__":
    path = r"\\AMNAS\amlex\Admin\Dorin Holban\DHO Time\DHO Time 2025\DHO Time - 2025-07.xlsx"
    entries = parse_DH_time_sheet(path)
    for e in entries:
        with open("entries.txt", "a", encoding="utf-8") as f:
            f.write(str(e) + "\n")

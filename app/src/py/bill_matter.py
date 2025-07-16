from base import *
from parse_json import *

def main():
    win = connect_to_pclaw()
    win.set_focus()

    data = read_stdin_json()
    matter = get_matter(data)

    register_matter(matter)
    sleep(3)

    date = ocr_get_latest_date()

    if date:
        bill_matter(matter, date)

if __name__ == "__main__":
    main()

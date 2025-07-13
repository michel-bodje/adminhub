import sys
from base import *
from parse_json import *

def main():
    win = connect_to_pclaw()
    win.set_focus()

    # Get matter from command line argument
    if len(sys.argv) > 1:
        matter = sys.argv[1]
    else:
        matter = ""

    register_matter(matter)
    sleep(3)

    date = ocr_get_latest_date()

    if date:
        bill_matter(matter, date)

if __name__ == "__main__":
    main()

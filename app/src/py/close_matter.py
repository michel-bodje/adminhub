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

    close_matter(matter)

if __name__ == "__main__":
    main()

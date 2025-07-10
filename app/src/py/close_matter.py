from base import *
from parse_json import *

def main():
    win = connect_to_pclaw()
    win.set_focus()

    # Matter string needs to be passed through JSON or other means
    matter = ""  # TODO: Load matter value appropriately

    close_matter(matter)

if __name__ == "__main__":
    main()

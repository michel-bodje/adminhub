from pclaw import *
from parse_json import *

def main():
    _, data = startup()
    matter = get_matter(data)

    close_matter(matter)

if __name__ == "__main__":
    main()
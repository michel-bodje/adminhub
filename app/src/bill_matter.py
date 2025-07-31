from pclaw import *
from parse_json import *

def main():
    app, data = startup()
    matter = get_matter(data)

    bill_matter(app, matter)

if __name__ == "__main__":
    main()
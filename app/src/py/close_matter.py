from pclaw import *
from parse_json import *

def main():
    app = connect_to_pclaw()
    app.set_focus()

    data = read_stdin_json()
    matter = get_matter(data)

    close_matter(matter)

if __name__ == "__main__":
    main()

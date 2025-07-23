from pclaw import *
from parse_json import *

def startup():
    """ Connects to PCLaw and sets focus. """
    app = connect_to_pclaw()
    app.set_focus()
    data = read_stdin_json()
    return app, data

def main():
    app, data = startup()
    matter = get_matter(data)

    bill_matter(app, matter)

if __name__ == "__main__":
    main()

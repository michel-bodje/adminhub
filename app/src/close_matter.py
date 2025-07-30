from pclaw import *
from parse_json import *

def startup():
    """ Connects to PCLaw and sets focus. """
    app = connect_to_pclaw()
    app.set_focus()
    data = read_json()
    sleep(3)  # Allow time for PCLaw to set focus
    return app, data

def main():
    _, data = startup()
    matter = get_matter(data)

    close_matter(matter)

if __name__ == "__main__":
    main()

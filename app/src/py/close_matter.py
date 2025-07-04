from base import *
from parse_json import *

win = connect_to_pclaw()
win.set_focus()

# matter string needs to be passed thru json
matter = ""

close_matter(matter)

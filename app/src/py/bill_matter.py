from base import *
from parse_json import *

win = connect_to_pclaw()
win.set_focus()

# matter string needs to be passed thru json
matter = ""

register_matter(matter)
sleep(3)

date = ocr_get_latest_date()

if (date):
    bill_matter(matter, date)

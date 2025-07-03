from base import *
from parse_json import *

main = connect_to_pclaw()
main.set_focus()

# replace with form data? how to pass strings?
matter = "4873-001"

register_matter(matter)
sleep(3)

date = ocr_get_latest_date_if_no_trust()

bill_matter(matter, date)
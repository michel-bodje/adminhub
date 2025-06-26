from base import *
from parse_json import *

main = connect_to_pclaw()
main.set_focus()

matter: str = "4849-001"

register_matter(matter)
sleep(5)
bill_matter(matter, "2025/6/6")
from pywinauto.application import Application
from pywinauto.findwindows import find_windows
from pywinauto.base_wrapper import BaseWrapper
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto import keyboard
from pywinauto.keyboard import send_keys
from time import sleep
from datetime import datetime
import os
import re
import pytesseract
import pyautogui

def connect_to_pclaw():
    """ Connects to the PCLaw Enterprise application and returns the main window. """
    hwnds = find_windows(title_re=".*PCLaw® Enterprise.*")
    app = Application(backend="uia").connect(handle=hwnds[0])
    main_win = app.window(handle=hwnds[0])
    return main_win

def get_dialog(parent_window, dialog_title):
    """ Returns a dialog window by title."""
    dlg = parent_window.child_window(title=dialog_title, control_type="Window")
    dlg.wait("visible enabled ready", timeout=15)
    return dlg

def new_matter_dialog():
    """ Opens the New Matter dialog using keyboard navigation. """
    send_keys('%f')
    sleep(0.5)
    send_keys('{DOWN}')
    sleep(0.5)
    send_keys('{RIGHT}')
    sleep(0.5)
    send_keys('{ENTER}')

def close_matter_dialog():
    """ Opens the Close Matter dialog using keyboard navigation. """
    send_keys('%f')
    sleep(0.5)
    send_keys('{DOWN}')
    sleep(0.5)
    send_keys('{RIGHT}')
    sleep(0.5)
    send_keys('{DOWN}{DOWN}')
    sleep(0.5)
    send_keys('{ENTER}')

def register_dialog():
    """ Opens the Register Matter dialog using keyboard navigation. """
    send_keys('%d')
    sleep(0.5)
    send_keys('{UP}')
    sleep(0.5)
    send_keys('{ENTER}')

def register_matter(matter_number: str):
    """ Opens the chosen matter in the Register. """
    register_dialog()
    send_keys("%m")
    send_keys(matter_number)
    sleep(0.5)
    send_keys("%s")
    sleep(0.5)
    send_keys("{ENTER}")

def bill_matter(matter_number: str, date: str = None, options: bool = False):
    """ Opens the Bill Matter dialog and fill with matter number and date. """
    if date is None:
        print("⚠️ No date provided for billing. Aborting.")
        return

    main = connect_to_pclaw()

    # Fill basic fields
    send_keys('^b')
    send_keys("%m")
    sleep(0.3)
    send_keys(matter_number)
    sleep(0.3)
    send_keys("%b")
    send_keys(date)
    sleep(0.3)
    send_keys("%a")
    send_keys(date)
    sleep(0.3)

    # Potential extra logic here: what if they want to use options?
    # If the dialog has options, we can handle them here.
    if(options):
        pass
    else:
        # Click OK
        main.child_window(title="OK", control_type="Button").click_input()

    # Wait for new dialog to appear
    sleep(10)
    # Validate default bill settings
    send_keys("%m")
    send_keys("{ENTER}")

    # Wait for the preview dialog to appear
    sleep(6)
    # print: send that shortcut enough times to trigger
    send_keys("%p")
    send_keys("%p")
    send_keys("%p")
    send_keys("%p")
    send_keys("%p")

    # Wait for the print to complete
    sleep(2)
    send_keys("{ENTER}")
    # Refresh
    sleep(0.3)
    send_keys("%m")
    send_keys("%s")
    sleep(0.3)
    send_keys("{ENTER}")

def close_matter(matter_number: str):
    """
    Opens the Close Matter dialog for the chosen matter.
    Also checks visually if closable or not
    """
    main = connect_to_pclaw()
    main.set_focus()

    close_matter_dialog()

    # Fill matter number
    send_keys("%m")
    send_keys(matter_number)
    send_keys("{TAB}")

    # Let PCLaw load like the atrocious software it is
    sleep(33)
    send_keys("{ENTER}")

    # Fill window, assume "No physical file"
    send_keys("d")
    send_keys("f")
    sleep(0.5)
    send_keys("No physical file", with_spaces=True)
    sleep(0.5)
    send_keys("{TAB}")
    send_keys("v")

    # Run the OCR script to check balances
    # before closing, we need to confirm balances are zero
    no_balance = ocr_get_balance()

    if no_balance:
        sleep(0.5)
        send_keys("{ENTER}")
        sleep(0.5)
        send_keys("{ENTER}")
        sleep(0.5)
        send_keys("{ENTER}")
    else:
        print("⚠️ Remaining balance is not zero. Matter should not be closed.")
        #ctypes.windll.user32.MessageBoxW(0, "⚠️ Remaining balance is not zero. Matter should not be closed until all balances are cleared.", "Warning", 0x30)

def ocr_get_balance():
    """ Uses OCR to extract financial data from the Close Matter dialog."""
    pytesseract.pytesseract.tesseract_cmd = os.path.join(os.path.dirname(__file__), "tesseract", "tesseract.exe")

    LABELS = ["Unbd D", "A/R", "Gen Rtnr", "Trust"]

    # === Find the Close Matter Window ===
    try:
        close_win = get_dialog(connect_to_pclaw(), "Close Matter")
        close_win.set_focus()
    except Exception as e:
        print("❌ Failed to locate Close Matter window:", e)
        return False

    # Get window bounds
    rect = close_win.rectangle()
    left, top, right, bottom = rect.left, rect.top, rect.right, rect.bottom
    width, height = right - left, bottom - top

    # Crop region = bottom ~25% of window where the financial data lives
    crop_left = int(left + width * 0.05)
    crop_right = int(left + width * 0.95)
    crop_top = int(top + height * 0.70)
    crop_bottom = int(top + height * 0.98)
    crop_width = crop_right - crop_left
    crop_height = crop_bottom - crop_top

    # Take screenshot of the cropped region
    screenshot = pyautogui.screenshot(region=(crop_left, crop_top, crop_width, crop_height))

    # (Optional) Save image to inspect
    # screenshot.save("debug_crop.png")

    # === OCR ===
    custom_config = r'--oem 3 --psm 6'
    text = pytesseract.image_to_string(screenshot, config=custom_config)
    # print("==== OCR Text ====")
    # print(text)

    # Extract function
    def extract_label_value(label, text):
        if label == "A/R":
            label_regex = r"A\S?R"
        elif label == "Gen Rtnr":
            label_regex = r"Gen\S*"
        else:
            label_regex = re.escape(label)

        # Search for the label
        label_match = re.search(label_regex, text, re.IGNORECASE)
        if not label_match:
            return None

        # Search for the first number AFTER the label
        post_label_text = text[label_match.end():]
        number_match = re.search(r"(-?\d+(?:\.\d{1,2})?)", post_label_text)
        if number_match:
            return float(number_match.group(1))

        return None

    # Final result dictionary
    amounts = {label: extract_label_value(label, text) for label in LABELS}

    # === Check and Display ===
    all_zero = True
    print("==== Extracted Amounts ====")
    for label, value in amounts.items():
        if value is None:
            print(f"{label}: Not found")
            all_zero = False
        else:
            print(f"{label}: {value:.2f}")
            if value != 0:
                all_zero = False

    if all_zero:
        print("\n✅ All values are zero. Proceed.")
        return True
    else:
        print("\n❌ Not all values are zero. Abort.")
        return False

def ocr_get_latest_date():
    """
    Uses OCR to determine latest date on the Register,
    aborting if there is a trust balance.
    You need manual user checking when it comes to trusts.
    """
    pytesseract.pytesseract.tesseract_cmd = os.path.join(os.path.dirname(__file__), "tesseract", "tesseract.exe")

    # === Locate Register window ===
    try:
        register_win = get_dialog(connect_to_pclaw(), "Register...")
        register_win.set_focus()
    except Exception as e:
        print("❌ Failed to locate Register window:", e)
        return None

    # === Screenshot bottom-left where "Trust" balance appears ===
    rect = register_win.rectangle()
    left, top, right, bottom = rect.left, rect.top, rect.right, rect.bottom
    width, height = right - left, bottom - top

    # Take full Register window screenshot
    full_screenshot = pyautogui.screenshot(region=(left, top, width, height))

    # Crop
    trust_crop = full_screenshot.crop((
        0,                      # start from left
        int(height * 0.75),     # lower 25%
        int(width * 0.5),       # to middle
        height                  # to bottom
    ))

    # Save for debugging
    # trust_crop.save("debug_trust_crop.png")

    trust_text = pytesseract.image_to_string(trust_crop, config='--oem 3 --psm 6')
    # print("==== TRUST TEXT ====")
    # print(trust_text)

    # Match for "Trust"
    trust_match = re.search(r"Trust[:\s]*(-?\d+(?:\.\d{1,2})?)", trust_text, re.IGNORECASE)
    trust_val = float(trust_match.group(1)) if trust_match else None

    # Match for "Gen Rtnr"
    retainer_match = re.search(r"Gen(?:\s+\S+)?[:\s]*(-?\d+(?:\.\d{1,2})?)", trust_text, re.IGNORECASE)
    retainer_val = float(retainer_match.group(1)) if retainer_match else None

    print("Trust Balance:", trust_val)
    print("Gen Rtnr :", retainer_val)

    if trust_val is None or trust_val > 0:
        print("❌ Trust balance is not zero. Abort.")
        return None
    
    if retainer_val is None or retainer_val not in (143.72, 402.41):
        print(f"❌ Gen Rtnr is {retainer_val}, requiring manual verification. Abort.")
        return None

    # === Screenshot top half (table rows) ===

    # Just upper half, starting from left
    table_crop = full_screenshot.crop((
        0, 0,
        int(width * 0.5), int(height * 0.5)
    ))

    # Save for debugging
    # table_crop.save("debug_table_crop.png")

    table_text = pytesseract.image_to_string(table_crop, config='--oem 3 --psm 6')
    # print("==== TABLE TEXT ====")
    # print(table_text)

    # === Extract dates ===
    date_matches = re.findall(r"\b(20\d{2})[/-](\d{1,2})[/-](\d{1,2})\b", table_text)
    dates = []
    for y, m, d in date_matches:
        try:
            dates.append(datetime(int(y), int(m), int(d)))
        except:
            continue

    if not dates:
        print("❌ No valid dates found.")
        return None

    latest = max(dates)
    print("✅ Latest Date:", latest.strftime("%Y/%m/%d"))
    return latest.strftime("%Y/%m/%d")

def find_nearest_input(label: BaseWrapper, inputs: list[BaseWrapper]):
    label_rect = label.rectangle()
    label_mid_y = (label_rect.top + label_rect.bottom) // 2
    best_ctrl = None
    min_dist = float("inf")

    for ctrl in inputs:
        rect = ctrl.rectangle()
        mid_y = (rect.top + rect.bottom) // 2
        dx = abs(rect.left - label_rect.right)
        dy = abs(mid_y - label_mid_y)
        if dx < 300 and dy < 20:
            dist = dx + dy
            if dist < min_dist:
                min_dist = dist
                best_ctrl = ctrl

    return best_ctrl

def get_label_controls(parent_window, needed_labels: set[str]):
    return [
        ctrl for ctrl in parent_window.descendants(control_type="Text")
        if ctrl.window_text().strip() in needed_labels
    ]

def get_input_controls(parent_window):
    return parent_window.descendants(control_type="Edit") + parent_window.descendants(control_type="Document")

def build_label_input_map(parent_window, fields: dict[str, str]):
    needed_labels = set(fields.keys())
    labels = get_label_controls(parent_window, needed_labels)
    inputs = get_input_controls(parent_window)

    label_input_map = {}
    for label in labels:
        label_text = label.window_text().strip()
        input_field = find_nearest_input(label, inputs)
        if input_field:
            label_input_map[label_text] = input_field

    return label_input_map

def fill_form_fields(label_input_map: dict[str, BaseWrapper], fields: dict[str, str]):
    for label, value in fields.items():
        input_field = label_input_map.get(label)
        if input_field:
            input_field.set_edit_text(value)
            print(f"✓ Filled '{label}' with '{value}'")
        else:
            print(f"⚠️ Could not find input for '{label}'")

def send_ctrl_arrow(direction: str = "right"):
    """
    Sends Ctrl+Right or Ctrl+Left.
    direction: "right" or "left".
    """
    seq = "^({RIGHT})" if direction.lower() == "right" else "^({LEFT})"
    keyboard.send_keys(seq)
    sleep(0.5)

def clear_focus(dlg):
    """
    Clear focus so that Ctrl+Arrow works.
    Seems to break navigation?
    """
    try:
        dlg.set_focus()
        print("Focus cleared")
        sleep(0.5)
    except:
        print("Failed to clear focus")
        return False
    return True

def focus_first_edit(dlg):
    """
    Focuses the first editable input in the dialog so that Ctrl+Arrow works.
    """
    try:
        edits = dlg.descendants(control_type="Edit")
        if edits:
            # Use set_focus() if available; else click_input to focus
            try:
                edits[0].set_focus()
            except:
                edits[0].click_input()
            sleep(0.1)
            return True
    except Exception:
        pass
    return False

def move_tab(direction="right", repeat=1, dlg=None):
    """
    Moves tab left or right a number of times, focusing the first edit each time.
    direction: "left" or "right"
    repeat: number of times to move
    dlg: dialog window to operate on (required)
    """
    if dlg is None:
        raise ValueError("dlg parameter is required")
    for _ in range(repeat):
        focus_first_edit(dlg)
        send_ctrl_arrow(direction)

def detect_active_tab(dlg):
    """
    Returns "main", "billing", or "custom" by checking for a unique label in each pane.
    Must be called right after a switch so contents are loaded.
    """
    panes = dlg.descendants(control_type="Pane")
    # For each pane matching the main-content geometry, check for unique texts:
    for pane in panes:
        # Check boundingRect matches content area if needed, or just try:
        texts = [t.window_text().strip() for t in pane.descendants(control_type="Text") if t.window_text().strip()]
        # Look for Billing-unique:
        if any("Billing Template" in txt for txt in texts):
            return "billing"
        # Look for Custom-unique:
        if any("BARREAU-" in txt for txt in texts):
            return "custom"
        # Look for Main-unique: e.g. “Type of Law” and “Responsible Lawyer” appear in Main
        if "Type of Law" in texts and "Responsible Lawyer" in texts:
            return "main"
    return None

def go_to_main(dlg):
    """
    Navigates to Main tab from wherever.
    """
    move_tab("left", 2, dlg)

def go_to_billing(dlg):
    """
    From Main, goes to Billing.
    Call go_to_main first to ensure starting at Main.
    """
    move_tab("right", 1, dlg)  # Main -> Billing

def go_to_custom(dlg):
    """
    From Main, goes to Custom via two rights.
    """
    move_tab("right", 2, dlg)  # Main -> Custom

def click_billing_checkbox(dlg):
    """
    Finds and clicks the 'Allow Bill Setting Overrides' checkbox in the Billing pane.
    """
    checkbox = dlg.child_window(auto_id="1105", control_type="CheckBox")
    try:
        checkbox.click_input()
        print("✓ Clicked 'Allow Bill Setting Overrides' checkbox")
        return True
    except Exception as e:
        print(f"❌ Failed to click checkbox: {e}")
        return False
    """
    # Alternative method using label search:
    # 1. Find the label element for “Allow Bill Setting Overrides”
    try:
        label = dlg.child_window(title="Allow Bill Setting Overrides", control_type="DataItem")
    except ElementNotFoundError:
        print("❌ Label DataItem not found")
        return False

    parent_info = label.element_info.parent
    if not parent_info:
        print("❌ Label has no parent")
        return False

    try:
        parent_wrapper = UIAWrapper(parent_info)
    except Exception:
        print("❌ Cannot wrap parent")
        return False

    # 3. Among parent.children(), inspect nameless DataItem nodes for a CheckBox child
    checkbox = None
    for child in parent_wrapper.children():
        # We only inspect DataItem children (the nameless row container)
        # and skip those whose title is non-empty (to target the unnamed container)
        try:
            ctrl_type = child.element_info.control_type
            title = child.window_text().strip()
        except Exception:
            continue
        if ctrl_type.lower() == "dataitem" and title == "":
            # Look for a CheckBox descendant under this nameless DataItem
            try:
                cbs = child.descendants(control_type="CheckBox")
            except Exception:
                cbs = []
            if cbs:
                # Take the first CheckBox found
                checkbox = cbs[0]
                break
    """

def find_edit_by_value(dlg, target_value):
    """
    Search all Edit controls and return the first whose current value equals target_value.
    """
    for edit in dlg.descendants(control_type="Edit"):
        try:
            val = edit.get_value()  # requires UIA backend
        except Exception:
            print(f"⚠️ Failed to get value for Edit control: {edit}")
            continue
        if val == target_value:
            print(f"✓ Found Edit with value '{target_value}': {edit}")
            return edit
    return None

def find_edit_by_label(dlg, target_value):
    """Search for an Edit control under a nameless DataItem that has the target_value.
    If the main search fails, it falls back to searching by label.
    """
    print("⚠️ Falling back to label-based search")
    billing_pane = dlg.child_window(auto_id="cmainpage", control_type="Pane")
    parent = billing_pane.element_info.parent
    parent_wrapper = UIAWrapper(parent)
    for child in parent_wrapper.children():
        try:
            ctrl_type = child.element_info.control_type
            title = child.window_text().strip()
        except Exception:
            continue
        if ctrl_type.lower() == "pane" and title == "":
            # Look for an edit descendant under this nameless DataItem
            try:
                edits = child.descendants(control_type="Edit")
            except Exception:
                print(f"⚠️ Failed to get descendants for {child}")
                edits = []
            for edit in edits:
                try:
                    val = edit.get_value()
                except Exception:
                    print(f"⚠️ Failed to get value for Edit control: {edit}")
                    continue
                if val == target_value:
                    print(f"✓ Found Edit with value '{target_value}': {edit}")
                    return edit
    print(f"⚠️ No Edit found with value '{target_value}'")
    return None

def fill_billing_tab_by_override(dlg):
    """ Fills the Billing tab with the required values.
    This function assumes the Billing tab is already active.
    """
    # Use the robust checkbox clicker
    if not click_billing_checkbox(dlg):
        print("⚠️ Could not check 'Allow Bill Setting Overrides' via label/parent search")
        return

    old_value = "RABILEN5"
    new_value = "RABILFR5"

    edit = find_edit_by_value(dlg, old_value)
    if not edit:
        print(f"⚠️ No Edit found with value '{old_value}', trying label search")
        edit = find_edit_by_label(dlg, old_value)
    if edit:
        edit.set_edit_text(new_value)
        keyboard.send_keys("{TAB}")
        print("✓ Replaced Billing Template value")
        return True
    else:
        print(f"⚠️ Could not find Edit with value '{old_value}'")
        print("❌ Failed to fill Billing tab")
        return False

def fill_billing_tab_by_value(dlg):
    """
    Fills the Billing tab by finding the Edit with value 'Default' and replacing it with 'Facture francais'.
    """
    edit = find_edit_by_value(dlg, "Default")
    if edit:
        edit.set_edit_text("Facture francais")
        keyboard.send_keys(" ")
        keyboard.send_keys("{BACKSPACE}")
        print("✓ Replaced 'Default' with 'Facture francais' in Billing tab")
        return True
    else:
        print("⚠️ Could not find Edit with value 'Default'")
        return False

def fill_custom_tab(dlg):
    """ Fills all editable fields in the Custom tab with 'n/a'.
    """
    edits = dlg.descendants(control_type="Edit")
    filled = 0
    for ctrl in edits:
        try:
            if ctrl.is_enabled():
                ctrl.set_edit_text("n/a")
                filled += 1
        except:
            continue
    print(f"✓ Filled {filled} editable fields in Custom tab with 'n/a'")

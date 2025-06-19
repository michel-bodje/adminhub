from pywinauto.application import Application
from pywinauto.findwindows import find_windows
from pywinauto.base_wrapper import BaseWrapper
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto import keyboard
from time import sleep

def connect_to_pclaw():
    hwnds = find_windows(title_re=".*PCLaw® Enterprise.*")
    app = Application(backend="uia").connect(handle=hwnds[0])
    main_win = app.window(handle=hwnds[0])
    main_win.set_focus()
    return main_win

def open_dialog(parent_window, dialog_title, button_title):
    parent_window.child_window(title=button_title, control_type="Button").click_input()
    sleep(1)
    dlg = parent_window.child_window(title=dialog_title, control_type="Window")
    # dlg.wait("visible enabled ready")
    return dlg

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
    sleep(0.2)

def focus_first_edit(dlg):
    """
    Focuses the first editable input in the dialog so that Ctrl+Arrow works.
    Maybe just use tab instead?
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

def move_tab(direction="left", repeat=1, dlg=None):
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
                sleep(0.2)

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
    current = detect_active_tab(dlg)
    if current == "main":
        return True
    if current == "billing":
        move_tab("left", 1, dlg)
    elif current == "custom":       
        move_tab("left", 2, dlg)
    else:
        # Unknown state: try left twice to land on Main
        print("⚠️ Unknown tab state, trying to go to Main")
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
                
def fill_billing_tab(dlg):
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
        print("✓ Replaced Billing Template value")
        return True
    else:
        print(f"⚠️ Could not find Edit with value '{old_value}'")
        print("❌ Failed to fill Billing tab")
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

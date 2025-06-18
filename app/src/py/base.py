from pywinauto.application import Application
from pywinauto.findwindows import find_windows
from pywinauto.base_wrapper import BaseWrapper
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
    dlg.wait("visible enabled ready", timeout=3)
    dlg.set_focus()
    return dlg

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

def send_ctrl_arrow(direction: str = "right"):
    """
    Sends Ctrl+Right or Ctrl+Left.
    direction: "right" or "left".
    """
    seq = "^({RIGHT})" if direction.lower() == "right" else "^({LEFT})"
    keyboard.send_keys(seq)
    sleep(0.2)

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
    # Ensure starting at Main
    if not go_to_main(dlg):
        print("⚠️ Could not confirm Main before going to Billing")
    move_tab("right", 1, dlg)  # Main -> Billing
    return detect_active_tab(dlg) == "billing"

def go_to_custom(dlg):
    """
    From Main, goes to Custom via two rights.
    """
    if not go_to_main(dlg):
        print("⚠️ Could not confirm Main before going to Custom")
    move_tab("right", 2, dlg)  # Main -> Custom
    return detect_active_tab(dlg) == "custom"


def get_label_controls(parent_window, needed_labels: set[str]):
    return [
        ctrl for ctrl in parent_window.descendants(control_type="Text")
        if ctrl.window_text().strip() in needed_labels
    ]

def get_input_controls(parent_window):
    return parent_window.descendants(control_type="Edit") + parent_window.descendants(control_type="Document")

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

def fill_billing_tab(pane):
    try:
        checkbox = pane.child_window(title="Allow Bill Settings Overrides", control_type="CheckBox")
        if not checkbox.get_toggle_state():
            checkbox.toggle()
            print("✓ Checked 'Allow Bill Settings Overrides'")
    except Exception as e:
        print(f"⚠️ Billing checkbox error: {e}")

    try:
        billing_field = pane.child_window(title="Billing Template", control_type="Edit")
        billing_field.set_edit_text("RABILFR5")
        print("✓ Filled Billing Template with 'RABILFR5'")
    except Exception as e:
        print(f"⚠️ Billing field error: {e}")

def fill_custom_tab_fields(pane):
    edits = pane.descendants(control_type="Edit")
    filled = 0
    for ctrl in edits:
        try:
            if ctrl.is_enabled():
                ctrl.set_edit_text("n/a")
                filled += 1
        except:
            continue
    print(f"✓ Filled {filled} editable fields in Custom tab with 'n/a'")

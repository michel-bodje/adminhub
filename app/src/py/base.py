from pywinauto.application import Application
from pywinauto.findwindows import find_windows
from pywinauto.base_wrapper import BaseWrapper
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

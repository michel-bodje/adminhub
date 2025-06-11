import win32com.client
import os
import time

def import_macros_from_folder(folder_path):
    outlook = win32com.client.Dispatch("Outlook.Application")
    vbe = outlook.VBE
    project = vbe.ActiveVBProject

    # List all .bas files in the folder
    modules = [f for f in os.listdir(folder_path) if f.lower().endswith(".bas")]

    print(f"Found {len(modules)} VBA modules to import...")
    for mod in modules:
        full_path = os.path.join(folder_path, mod)
        try:
            print(f"Importing {mod}...")
            project.VBComponents.Import(full_path)
        except Exception as e:
            print(f"Error importing {mod}: {e}")

    print("✅ All modules imported into Outlook.")
    time.sleep(1)

if __name__ == "__main__":
    macro_folder = os.path.abspath(os.path.join(os.path.dirname(__file__), "office"))
    import_macros_from_folder(macro_folder)

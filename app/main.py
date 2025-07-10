import webview
import json
import os
import subprocess
import sys

# Get the root directory of the application
def get_base_path():
    if getattr(sys, 'frozen', False):  # Running as .exe (PyInstaller bundles it)
        return getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))  # temp folder with bundled files
    else:
        # Go up one directory from the current file's directory to get the project root
        return os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

ROOT_DIR = get_base_path()

WEB_DIR = os.path.join(ROOT_DIR, 'app', 'web')
SRC_DIR = os.path.join(ROOT_DIR, 'app', 'src')
DATA_JSON = os.path.join(ROOT_DIR, 'data', 'data.json')
INDEX_HTML = os.path.join(WEB_DIR, 'index.html')

class HubAPI:
    def submit_form(self, form, lwy, dets):
        try:
            data_path = DATA_JSON
            data = {
                "form": form,
                "lawyer": lwy,
                "caseDetails": dets
            }

            with open(data_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)

            print("[AdminHub] Form data saved to:", data_path)
            return { "message": "Form received and saved." }

        except Exception as e:
            print("[AdminHub] Error:", str(e))
            return { "message": "Failed to save form: " + str(e) }

    def get_lawyers(self):
        lawyers_path = os.path.join(WEB_DIR, "js", "lawyers.json")
        try:
            with open(lawyers_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data["lawyers"]
        except Exception as e:
            return {"error": str(e)}
        
    def run_exe(self, script_name):
        script_path = os.path.join(SRC_DIR, "cs", "bin", script_name + ".exe")
        try:
            # subprocess.run([script_path, ROOT_DIR], check=True)
            os.startfile(script_path)
            return {"message": f"Running EXE: {script_name}"}
        except Exception as e:
            print(f"[AdminHub] Error running EXE '{script_name}': {e}")
            return {"error": str(e)}

    def run_py(self, script_name):
        """
        In final version, use pyinstaller to bundle Python scripts into .exe files.
        For development, we can run the scripts directly with pythonw.
        """
        script_path = os.path.join(SRC_DIR, "py", script_name + ".py")
        try:
            subprocess.run(["pythonw", script_path, ROOT_DIR], check=True)
            return {"message": f"Running Python script: {script_name}"}
        except Exception as e:
            print(f"[AdminHub] Error running Python script '{script_name}': {e}")
            return {"error": str(e)}

if __name__ == '__main__':
    html_path = INDEX_HTML
    api = HubAPI()
    webview.create_window("Amlex Admin Hub", html_path, js_api=api, width=425, height=650)
    webview.start(debug=False, gui='edgechromium')

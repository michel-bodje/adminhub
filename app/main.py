import webview
import json
import os
import subprocess
import sys

# Get the root directory of the application
def get_base_path():
    if getattr(sys, 'frozen', False):  # Running as .exe (PyInstaller bundles it)
        return sys._MEIPASS  # temp folder with bundled files
    else:
        return os.path.abspath(os.path.join(__file__, ".."))  # Running in development mode

BASE_PATH = get_base_path()

WEB_DIR = os.path.join(BASE_PATH, 'web')
SRC_DIR = os.path.join(BASE_PATH, 'src')
DATA_JSON = os.path.join(BASE_PATH, 'data.json')
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
            return { "error": str(e) }
        
    def run_exe(self, script_name):
        script_path = os.path.join(SRC_DIR, "cs", "bin", script_name + ".exe")
        subprocess.run([script_path, BASE_PATH], check=True)

    def run_py(self, script_name):
        script_path = os.path.join(SRC_DIR, "py", script_name + ".py")
        subprocess.run(["pythonw", script_path, BASE_PATH], check=True)

if __name__ == '__main__':
    html_path = INDEX_HTML
    api = HubAPI()
    webview.create_window("Amlex Admin Hub", html_path, js_api=api, width=500, height=650)
    webview.start(debug=True, gui='edgechromium')

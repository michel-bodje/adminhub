import webview
import json
import os
import subprocess
import sys

# Get the root directory of the application
def get_base_path():
    if getattr(sys, 'frozen', False):
        # temp folder with bundled files
        return getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    else:
        # Go up one directory from the current file's directory to get the project root
        return os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

ROOT_DIR = get_base_path()

WEB_DIR = os.path.join(ROOT_DIR, 'app', 'web')
SRC_DIR = os.path.join(ROOT_DIR, 'app', 'src')
DATA_JSON = os.path.abspath(os.path.join(get_base_path(), 'data', 'data.json'))
INDEX_HTML = os.path.join(WEB_DIR, 'index.html')

class HubAPI:
    """ API for the Amlex Admin Hub.
    This class provides methods to interact with the hub, including submitting forms,
    retrieving lawyer data, and running scripts.
    """
    def _log(self, message):
        print(f"[AdminHub] {message}")

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

    def run(self, script_name):
        """
        Runs a script based on its extension.
        For .exe: runs the executable.
        For .py: in development, runs the script with pythonw.
        In final version, use pyinstaller to bundle Python scripts into .exe files.
        """
        # Determine extension if present, else default to .exe
        base, ext = os.path.splitext(script_name)
        if not ext:
            # Default to .exe if no extension provided
            ext = ".exe"
            script_name = script_name + ext
        if ext == ".exe":
            script_path = os.path.join(SRC_DIR, "cs", "bin", script_name)
            try:
                subprocess.run([script_path], check=True)
                self._log(f"Running EXE: {script_name}")
                return {"message": f"Running EXE: {script_name}"}
            except Exception as e:
                self._log(f"Error running EXE '{script_name}': {e}")
                return {"error": str(e)}
        elif ext == ".py":
            script_path = os.path.join(SRC_DIR, "py", script_name)
            try:
                subprocess.run(["pythonw", script_path], check=True)
                self._log(f"Running Python script: {script_name}")
                return {"message": f"Running Python script: {script_name}"}
            except Exception as e:
                self._log(f"Error running Python script '{script_name}': {e}")
                return {"error": str(e)}
        else:
            self._log(f"Unknown script extension: {ext}")
            return {"error": f"Unknown script extension: {ext}"}

def main():
    api = HubAPI()
    webview.create_window("Amlex Admin Hub", INDEX_HTML, js_api=api, width=425, height=650)
    webview.start(debug=False, gui='edgechromium')

if __name__ == '__main__':
    main()
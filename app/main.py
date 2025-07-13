import os
import json
import subprocess
import webview
from src.py.config import *

class HubAPI:
    """ API for the Amlex Admin Hub.
    This class provides methods to interact with the hub from JavaScript.
    """
    def _log(self, message):
        print(f"[AdminHub] {message}")

    def submit_form(self, form, lwy, dets):
        """ Saves client, lawyer and case data to a JSON file.
        """
        try:
            data_path = DATA_JSON
            data = {
                "form": form,
                "lawyer": lwy,
                "caseDetails": dets
            }
            with open(data_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
            self._log(f"Form data saved to: {data_path}")
            return { "message": "Form received and saved." }
        except Exception as e:
            self._log(f"Error: {str(e)}")
            return { "message": "Failed to save form: " + str(e) }

    def get_lawyers(self):
        """ Retrieves the list of lawyers from a JSON file.
        """
        lawyers_path = os.path.join(WEB_DIR, "js", "lawyers.json")
        try:
            with open(lawyers_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data["lawyers"]
        except Exception as e:
            return {"error": str(e)}

    def run(self, script_name, *args):
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
            script_path = os.path.join(BIN_DIR, script_name)
            try:
                subprocess.run([script_path] + list(args), check=True)
                self._log(f"Running EXE: {script_name} with args: {args}")
                return {"message": f"Running EXE: {script_name} with args: {args}"}
            except Exception as e:
                self._log(f"Error running EXE '{script_name}': {e}")
                return {"error": str(e)}
            
        elif ext == ".py":
            script_path = os.path.join(SRC_DIR, "py", script_name)
            try:
                subprocess.run(["pythonw", script_path] + list(args), check=True)
                self._log(f"Running Python script: {script_name} with args: {args}")
                return {"message": f"Running Python script: {script_name} with args: {args}"}
            except Exception as e:
                self._log(f"Error running Python script '{script_name}': {e}")
                return {"error": str(e)}
            
        else:
            self._log(f"Unknown script extension: {ext}")
            return {"error": f"Unknown script extension: {ext}"}

def main():
    """ Main function to start the webview application.
    """
    api = HubAPI()
    webview.create_window("Amlex Admin Hub", INDEX_HTML, js_api=api, width=425, height=650)
    webview.start(debug=False, gui='edgechromium')

if __name__ == '__main__':
    main()
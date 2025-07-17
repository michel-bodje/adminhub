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

    def format_form(self, form, case, lwy):
        """ Returns structured JSON. """
        try:
            data = {
                "form": form,
                "case": case,
                "lawyer": lwy,
            }
            return { "json": json.dumps(data, indent=2) }
        except Exception as e:
            self._log(f"Error formatting form data: {str(e)}")
            return { "error": str(e) }

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

    def run(self, exe_name, json_blob):
        """ Runs a specified executable with the provided JSON input. """
        self._log(f"run() called with script_name: {exe_name}")
        try:
            path = os.path.join(BIN_DIR, exe_name + ".exe")
            proc = subprocess.run(
                [path],
                input=json_blob.encode("utf-8"),
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=True
            )
            return { "output": proc.stdout.decode() }
        except Exception as e:
            return { "error": str(e) }

def main():
    """ Main function to start the webview application.
    """
    api = HubAPI()
    webview.create_window("Amlex Admin Hub", INDEX_HTML, js_api=api, width=425, height=650)
    webview.start(debug=True, gui='edgechromium')

if __name__ == '__main__':
    main()
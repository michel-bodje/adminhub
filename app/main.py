from src.py.config import *
import json
import subprocess
import webview

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

    def run(self, script_name, json_blob):
        """ Runs a specified script with the provided JSON input. """
        self._log(f"run() called with script_name: {script_name}")
        # self._log(f"run() called with json_blob: {json_blob}")

        try:
            path = os.path.join(SRC_DIR, "py", script_name + ".py")
            if not os.path.exists(path):
                return { "error": f"Script not found: {path}" }
            
            # Call it via the Python interpreter
            proc = subprocess.run(
                [sys.executable, path],
                input=json_blob.encode("utf-8"),
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=True
            )

            output = proc.stdout.decode("utf-8").strip()
            return { "output": output }

        except subprocess.CalledProcessError as e:
            # Subprocess ran but returned non-zero (e.g., raised exception)
            error_output = e.stderr.decode("utf-8").strip()
            return { "error": f"Subprocess error:\n{error_output}" }

        except Exception as e:
            # Anything else (path issue, bad encoding, etc.)
            return { "error": str(e) }

def main():
    """ Main function to start the webview application.
    """
    api = HubAPI()
    webview.create_window(
        "Amlex Admin Hub",
        INDEX_HTML,
        js_api=api,
        width=425,
        height=700,
        x = 875,
        y = 0)
    webview.start(debug=True, gui='edgechromium')

if __name__ == '__main__':
    main()
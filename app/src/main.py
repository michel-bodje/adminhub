from config import *
import json
import subprocess
import webview

class HubAPI:
    """ API for the Amlex Admin Hub.
    This class provides methods to interact with the hub from JavaScript.
    """
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
            log(f"Error formatting form data: {str(e)}")
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
        
    def select_timesheet_file(self):
        """Open a file dialog to select timesheet file and return the full path."""
        try:
            import tkinter as tk
            from tkinter import filedialog
            
            root = tk.Tk()
            root.withdraw()  # Hide the main window
            root.attributes('-topmost', True)  # Bring dialog to front
            
            file_path = filedialog.askopenfilename(
                title="Select Timesheet File",
                filetypes=[
                    ("Excel files", "*.xlsx *.xls"),
                    ("All files", "*.*")
                ],
                initialdir=r"\\AMNAS\amlex\Admin"  # Start in your admin directory
            )
            
            root.destroy()
            return file_path if file_path else None
            
        except Exception as e:
            print(f"Error in file selection: {e}")
            return None    
    
    def run(self, script_name, json_blob):
        """ Runs a specified script with the provided JSON input. 
        Returns parsed JSON if the script outputs JSON, otherwise returns text output.
        """
        log(f"run() called with script_name: {script_name}")
    
        try:
            path = os.path.join(SRC_DIR, script_name + ".py")
            if not os.path.exists(path):
                return { "error": f"Script not found: {path}" }
            
            proc = subprocess.run(
                [sys.executable, path],
                input=json_blob.encode("utf-8"),
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=True
            )
    
            output = proc.stdout.decode("utf-8").strip()
            
            # Try to parse as JSON - look for the last complete JSON object in output
            try:
                # If there are multiple lines, try to find the JSON line
                lines = output.split('\n')
                json_line = None
                
                # Look for a line that starts with { or [
                for line in reversed(lines):  # Start from the end
                    line = line.strip()
                    if line.startswith('{') or line.startswith('['):
                        json_line = line
                        break
                    
                if json_line:
                    return json.loads(json_line)
                else:
                    # Try parsing the whole output
                    return json.loads(output)
                    
            except json.JSONDecodeError:
                # Not JSON, return as text
                return { "output": output }
    
        except subprocess.CalledProcessError as e:
            error_output = e.stderr.decode("utf-8").strip()
            return { "error": f"Script failed: {error_output}" }
        except Exception as e:
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
    webview.start(debug=False, gui='edgechromium')

if __name__ == '__main__':
    main()
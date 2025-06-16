import webview
import json
import os
import subprocess

root_dir = os.path.abspath(os.path.join(__file__, "../../../"))

class LawHubAPI:
    def submit_form(self, form, lwy, dets):
        try:
            data_path = os.path.join(root_dir, "app", "data.json")
            data = {
                "form": form,
                "lawyer": lwy,
                "caseDetails": dets
            }

            with open(data_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)

            print("[LawHub] Form data saved to:", data_path)

            return { "message": "Form received and saved." }

        except Exception as e:
            print("[LawHub] Error:", str(e))
            return { "message": "Failed to save form: " + str(e) }

    def run_script(self, script_name):
        script_path = os.path.join(root_dir, "app", "office", "bin", script_name + ".exe")
        subprocess.run([script_path], check=True)
        
    def get_lawyers(self):
        lawyers_path = os.path.join(root_dir, 'app', "web", "js", "lawyers.json")
        try:
            with open(lawyers_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data["lawyers"]
        except Exception as e:
            return { "error": str(e) }

if __name__ == '__main__':
    html_path = os.path.join(root_dir, "app", "web", "index.html")
    api = LawHubAPI()
    webview.create_window("LawHub", html_path, js_api=api, width=500, height=600)
    webview.start(debug=True, gui='edgechromium')

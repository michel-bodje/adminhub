import webview
import json
import os
import subprocess

class LawHubAPI:
    root_dir = os.path.abspath(os.path.join(__file__, "../../../"))

    def submit_form(self, form, lwy):
        try:
            json_path = os.path.join(self.root_dir, "app", "data.json")
            data = {
                "form": form,
                "lawyer": lwy
            }

            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)

            print("[LawHub] Form data saved to:", json_path)

            return { "message": "Form received and saved." }

        except Exception as e:
            print("[LawHub] Error:", str(e))
            return { "message": "Failed to save form: " + str(e) }

    def run_script(self, script_name):
        script_path = os.path.join(self.root_dir, "app", "office", script_name + ".exe")
        subprocess.run([script_path], check=True)
        
    def get_lawyers(self):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        path = os.path.join(base_dir, 'js', 'lawyers.json')
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data["lawyers"]
        except Exception as e:
            return { "error": str(e) }

if __name__ == '__main__':
    html_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "hub.html"))
    api = LawHubAPI()
    webview.create_window("LawHub", html_path, js_api=api, width=500, height=600)
    webview.start(debug=False, gui='edgechromium')

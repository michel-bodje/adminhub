import webview
import json
import os

class LawHubAPI:
    def submit_form(self, data):
        try:
            PROJECT_ROOT = os.path.abspath(os.path.join(__file__, "../../../"))
            json_path = os.path.join(PROJECT_ROOT, "data.json")

            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)

            print("[LawHub] Form data saved to:", json_path)

            return { "message": "Form received and saved." }

        except Exception as e:
            print("[LawHub] Error:", str(e))
            return { "message": "Failed to save form: " + str(e) }
        
    def get_lawyers(self):
        try:
            with open("js/lawyers.json", "r", encoding="utf-8") as f:
                data = json.load(f)
                return data["lawyers"]
        except Exception as e:
            return { "error": str(e) }

if __name__ == '__main__':
    html_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "hub.html"))
    api = LawHubAPI()
    webview.create_window("LawHub", html_path, js_api=api, width=500, height=740)
    webview.start(debug=True, gui='edgechromium')

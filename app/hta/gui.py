import webview
import json

class FormAPI:
    def handle_form(self, data):
        print("Received form:", data)

        # Optional: write to file
        with open("C:/LawHub/config/form.json", "w") as f:
            json.dump(data, f, indent=2)

        # Optional: trigger further automation
        # subprocess.run(["pythonw", "scripts/handle_form.py"])

        return {"message": "Form submitted successfully."}


if __name__ == "__main__":
    api = FormAPI()
    webview.create_window("LawHub Form", "form.html", js_api=api)
    webview.start()

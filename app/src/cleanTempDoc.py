import os
import sys
import time

def is_file_unlocked(path):
    try:
        with open(path, 'a+'):
            return True
    except OSError:
        return False

def cleanup():
    if len(sys.argv) != 2:
        print("Usage: cleanTempDoc.py <path>")
        return

    file_path = sys.argv[1]
    print(f"Watching temp file for release: {file_path}")

    timeout_minutes = 30
    check_interval = 30  # seconds
    waited = 0

    while waited < timeout_minutes * 60:
        if not os.path.exists(file_path):
            print("File already deleted.")
            return

        if is_file_unlocked(file_path):
            try:
                os.remove(file_path)
                print("Temp file deleted successfully.")
                return
            except Exception as e:
                print(f"Failed to delete file: {e}")

        time.sleep(check_interval)
        waited += check_interval

    print("Timeout reached. File still in use or deletion failed.")

if __name__ == "__main__":
    cleanup()

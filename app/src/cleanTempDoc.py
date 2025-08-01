import os
import sys
import time
import threading
from config import log

def is_file_unlocked(path):
    """Check if a file is no longer locked by another process."""
    try:
        with open(path, 'a+'):
            return True
    except OSError:
        return False

def process_clean_temp_doc(file_path, timeout_minutes=30, check_interval=30):
    """
    Clean up a temporary document file once it's no longer in use.
    
    This function can be called directly as a module function, avoiding
    the subprocess issues that occur with PyInstaller bundled executables.
    
    Args:
        file_path (str): Path to the temporary file to clean up
        timeout_minutes (int): Maximum time to wait before giving up
        check_interval (int): How often to check if file is available (seconds)
    Returns:
        dict: Result of the cleanup operation
    """
    log(f"Starting cleanup watch for temp file: {file_path}")
    
    # Validate that we actually have a file to clean up
    if not file_path:
        return {"error": "No file path provided for cleanup"}
    
    total_wait_time = timeout_minutes * 60
    waited = 0

    while waited < total_wait_time:
        # Check if file was already deleted by another process
        if not os.path.exists(file_path):
            log("Temp file was already deleted by another process.")
            return {"status": "success", "message": "File already deleted"}

        # Check if we can now access the file (meaning Word/Office has released it)
        if is_file_unlocked(file_path):
            try:
                os.remove(file_path)
                log("Temp file deleted successfully.")
                return {
                    "status": "success", 
                    "message": "Temp file deleted successfully",
                    "waited_seconds": waited
                }
            except Exception as e:
                error_msg = f"Failed to delete file: {e}"
                log(error_msg)
                return {"error": error_msg}

        # Wait before checking again
        time.sleep(check_interval)
        waited += check_interval
        
        # Log progress periodically so we know it's still working
        if waited % 300 == 0:  # Every 5 minutes
            log(f"Still waiting for file release... {waited//60} minutes elapsed")

    # Timeout reached
    timeout_msg = f"Timeout reached after {timeout_minutes} minutes. File still in use or deletion failed."
    log(timeout_msg)
    return {"error": timeout_msg}

def cleanup_temp_doc_async(file_path, timeout_minutes=30):
    """
    Start the cleanup process in a background thread to avoid blocking the main application.
    
    This is the function you'll call from office_utils.py instead of subprocess.
    
    Args:
        file_path (str): Path to the temporary file to clean up
        timeout_minutes (int): Maximum time to wait before giving up
    """
    def cleanup_worker():
        """Worker function that runs in the background thread."""
        try:
            result = process_clean_temp_doc(file_path, timeout_minutes)
            if result.get("error"):
                log(f"Cleanup error: {result['error']}")
            else:
                log(f"Cleanup completed: {result['message']}")
        except Exception as e:
            log(f"Unexpected error in cleanup worker: {e}")
    
    # Start the cleanup in a daemon thread so it doesn't prevent app shutdown
    cleanup_thread = threading.Thread(target=cleanup_worker, daemon=True)
    cleanup_thread.start()
    log(f"Started background cleanup thread for: {file_path}")

# Keep the original standalone functionality for backward compatibility and testing
def main():
    """Original standalone function for command-line usage."""
    if len(sys.argv) != 2:
        print("Usage: cleanTempDoc.py <path>")
        return

    file_path = sys.argv[1]
    result = process_clean_temp_doc(file_path)
    
    if result.get("error"):
        print(f"Error: {result['error']}")
        sys.exit(1)
    else:
        print(result["message"])

if __name__ == "__main__":
    main()
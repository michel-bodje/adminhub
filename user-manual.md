# AdminHub User Manual

## Overview

**AdminHub** is a Windows-based administrative tool designed to streamline and automate various office, legal, and document management tasks. It provides a suite of features for handling documents, emails, scheduling, billing, and more, with support for both English and French templates. The application is distributed as a standalone executable (`AdminHub.exe`) and can be run directly on Windows systems.

## Main Features

### 1. Document Generation and Management
- **Templates**: Supports Word (`.docx`) and HTML templates for contracts, receipts, feedback, office communications, and more, available in both English and French (`/templates/en/` and `/templates/fr/`).
- **Automated Document Creation**: Scripts like `wordContract.py` and `wordReceipt.py` automate the generation of legal documents and receipts.
- **Temporary Document Cleanup**: `cleanTempDoc.py` helps manage and remove temporary files.

### 2. Email Automation
- **Email Modules**: Automate sending confirmations, contracts, follow-ups, replies, and reviews using scripts such as `emailConfirmation.py`, `emailContract.py`, `emailFollowup.py`, `emailReply.py`, and `emailReview.py`.
- **Template-Based Emails**: HTML templates for various email types are provided for consistent communication.

### 3. Matter and Billing Management
- **Matter Handling**: Scripts like `new_matter.py` and `close_matter.py` help open and close legal matters efficiently.
- **Billing**: `bill_matter.py` and `time_entries.py` manage billing and time entry processes.
- **PC Law Integration**: `pclaw.py` provides integration with PC Law for legal billing and case management.

### 4. Scheduling and Office Utilities
- **Scheduler**: `scheduler.py` automates and manages scheduling appointments.
- **Office Utilities**: `office_utils.py` provides various helper functions for office administration.

### 5. OCR and Data Parsing
- **Tesseract Integration**: Bundled Tesseract OCR (`/tesseract/`) enables text extraction from images and scanned documents.
- **Data Parsing**: `parse_json.py` and `parse_timesheets.py` handle structured data extraction and timesheet processing.

### 6. Web Interface
- **Web Frontend**: The `/web/` directory contains a simple web interface with HTML, CSS, and JavaScript for user interaction.
- **Lawyer Management**: `lawyer.js` and `lawyers.json` manage lawyer profiles and related data.

## Getting Started

### Running the Application
- Launch `AdminHub.exe` directly, or use the provided batch scripts (`build.bat`, `launch.bat`) in the `/app/` directory.
- The application will present a GUI or web interface for user interaction.

### Configuration
- Configuration settings are managed in `config.py`.
- Templates can be customized in the `/templates/` directory.

### Building and Deployment
- Use `build.bat` to build the application from source if needed.
- The `/app/build/` directory contains build artifacts and logs.

## Directory Structure
- `/app/` – Build scripts, specs, and build outputs
- `/src/` – Main Python source code
- `/templates/` – Document and email templates (English and French)
- `/web/` – Web interface assets (HTML, CSS, JS, images)
- `/tesseract/` – OCR engine and data files

## Extending the Application
- Add new document or email templates in the appropriate `/templates/` subfolder.
- Implement new features by adding Python scripts in `/src/` and updating the module registry (`module_registry.py`).
- Update the web interface by editing files in `/web/`.

## Troubleshooting
- Check the `/app/build/` directory for logs and error reports.
- Ensure all dependencies listed in `requirements.txt` are installed.
- For OCR issues, verify the presence of `tesseract.exe` and the required language data in `/tesseract/tessdata/`.

## Credits
Developed and maintained by michel-bodje.

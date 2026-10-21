# Google Drive Batch Downloader & Document Pipeline

A Python automation pipeline to batch search, download, and convert documents from Google Drive into editable DOCX and standalone HTML, coupled with a browser automation script for field verification forms.

---

## Features

- **Batch Drive Search**: Queries Google Drive (including Shared Drives) by ID lists or console input.
- **Multi-Format Conversion**: Converts PDF files to DOCX (`pdf2docx`), and exports DOCX to HTML using Microsoft Word COM automation (`win32com`).
- **Idempotent Execution**: Automatically skips already processed IDs to save bandwidth and compute time.
- **Process Recovery**: Detects Word COM errors, closes hanging processes, and retries conversions automatically.
- **Failure Logging**: Tracks missing or failed IDs in `failed.txt` for easy auditing.
- **Browser Form Automation**: Companion script (`Script(Console).txt`) to auto-fill verification web forms and upload images under size limits (<500 KB).

---

## Architecture

```mermaid
flowchart LR
    A[check_ids.txt] --> B[drive_downloader.py]
    B --> C[Google Drive API v3]
    C --> D[PDF Download]
    D --> E[pdf2docx]
    E --> F[DOCX]
    F --> G[Word COM]
    G --> H[HTML Output]
```

---

## Tech Stack

- **Language**: Python 3.8+, JavaScript (ES6+)
- **APIs & Libraries**: Google Drive API v3, `google-auth-oauthlib`, `pdf2docx`, `pywin32`
- **Platform**: Windows 10/11 (required for Microsoft Word COM API)

---

## Directory Structure

```plaintext
drive-downloader/
├── drive_downloader.py          # Main batch download & conversion pipeline
├── Script(Console).txt          # Browser console automation script
├── check_ids.txt                # List of target IDs (one per line)
├── check_ids.example.txt        # Sample IDs file
├── requirements.txt             # Python dependencies
├── credentials.json.example     # Google OAuth credentials template
├── .gitignore                   # Excludes tokens, cache, and output folders
└── README.md
```

---

## Installation & Setup

### 1. Clone the Repository
```bash
git clone https://github.com/Ashesh88/drive-downloader.git
cd drive-downloader
```

### 2. Install Dependencies
```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Setup Google Drive API
1. Go to [Google Cloud Console](https://console.cloud.google.com/) and create a project.
2. Enable the **Google Drive API**.
3. Create an **OAuth 2.0 Client ID** (Application type: **Desktop App**).
4. Download the JSON file, rename it to `credentials.json`, and place it in the project root directory.

---

## Usage

### 1. Run Document Pipeline
1. Add target IDs to `check_ids.txt` (or paste when prompted):
   ```text
   169054659
   165575359
   151191894
   ```
2. Run the script:
   ```bash
   python drive_downloader.py
   ```
3. Outputs are generated in:
   - `downloads/` — Downloaded PDFs
   - `docx/` — Converted Word documents
   - `html/` — Final HTML files
   - `failed.txt` — Log of missing or failed IDs

### 2. Browser Verification Script
1. Open the target verification portal in your browser.
2. Press `F12` to open Developer Tools and go to the **Console** tab.
3. Paste the contents of `Script(Console).txt` and press `Enter` to auto-fill form fields and validate image uploads.

---

## License

This project is licensed under the [MIT License](LICENSE).

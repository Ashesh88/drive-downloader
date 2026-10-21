import io
import os
import re
import sys
import time

import win32com.client
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from pdf2docx import Converter

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
DOWNLOAD_DIR = "downloads"
DOCX_DIR = "docx"
HTML_DIR = "html"
IDS_FILE = "check_ids.txt"
FAILED_FILE = "failed.txt"

WD_FORMAT_HTML = 8  # Word "Web Page (*.htm, *.html)"


def get_service():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as f:
            f.write(creds.to_json())
    return build("drive", "v3", credentials=creds)


def get_word():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    return word


def search_file(service, check_id):
    safe_id = check_id.replace("'", "").replace("\\", "")
    query = f"name contains '{safe_id}' and trashed = false"
    res = service.files().list(
        q=query,
        fields="files(id, name, mimeType, modifiedTime)",
        orderBy="modifiedTime desc",
        pageSize=20,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = res.get("files", [])
    if not files:
        return None

    for f in files:
        if f["name"].lower() == f"{check_id}.pdf".lower():
            return f
    for f in files:
        if f["name"].lower().endswith(".pdf"):
            return f
    return files[0]


def download_file(service, file):
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    name = file["name"]
    mime = file["mimeType"]

    if mime.startswith("application/vnd.google-apps"):
        request = service.files().export_media(fileId=file["id"], mimeType="application/pdf")
        if not name.lower().endswith(".pdf"):
            name += ".pdf"
    else:
        request = service.files().get_media(fileId=file["id"], supportsAllDrives=True)

    path = os.path.join(DOWNLOAD_DIR, name)
    buf = io.FileIO(path, "wb")
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.close()
    return path


def pdf_to_docx(pdf_path):
    os.makedirs(DOCX_DIR, exist_ok=True)
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    docx_path = os.path.join(DOCX_DIR, base + ".docx")
    cv = Converter(pdf_path)
    try:
        cv.convert(docx_path)
    finally:
        cv.close()
    return docx_path


def docx_to_html(word, docx_path):
    os.makedirs(HTML_DIR, exist_ok=True)
    base = os.path.splitext(os.path.basename(docx_path))[0]
    src = os.path.abspath(docx_path)
    dst = os.path.abspath(os.path.join(HTML_DIR, base + ".htm"))
    doc = word.Documents.Open(src, ReadOnly=True)
    try:
        doc.SaveAs2(dst, FileFormat=WD_FORMAT_HTML)
    finally:
        doc.Close(False)
    return dst


def read_ids():
    raw = ""
    if os.path.exists(IDS_FILE):
        with open(IDS_FILE, "r", encoding="utf-8") as f:
            raw = f.read()
    if not raw.strip():
        print(f"'{IDS_FILE}' khali hai ya nahi mili. IDs yahan paste karo (comma/space se alag):")
        raw = input("IDs: ")
    ids = [x for x in re.split(r"[,\s;]+", raw) if x]
    # duplicates hatao, order same rakho
    return list(dict.fromkeys(ids))


def main():
    ids = read_ids()
    if not ids:
        print("Koi ID nahi mili.")
        sys.exit(0)

    service = get_service()
    word = get_word()

    total = len(ids)
    ok, skipped, not_found, failed = [], [], [], []
    start = time.time()
    print(f"\nTotal IDs: {total}\n")

    try:
        for i, check_id in enumerate(ids, 1):
            tag = f"[{i}/{total}] {check_id}"

            # pehle se ho chuki ho to skip
            if os.path.exists(os.path.join(HTML_DIR, check_id + ".htm")):
                print(f"{tag} -> SKIP (html already hai)")
                skipped.append(check_id)
                continue

            try:
                file = search_file(service, check_id)
                if not file:
                    print(f"{tag} -> NOT FOUND")
                    not_found.append(check_id)
                    continue

                pdf_path = download_file(service, file)
                docx_path = pdf_to_docx(pdf_path)

                try:
                    html_path = docx_to_html(word, docx_path)
                except Exception:
                    # Word atak gaya ho to restart karke ek baar aur try
                    try:
                        word.Quit()
                    except Exception:
                        pass
                    word = get_word()
                    html_path = docx_to_html(word, docx_path)

                print(f"{tag} -> OK ({os.path.basename(html_path)})")
                ok.append(check_id)
            except Exception as e:
                print(f"{tag} -> ERROR: {e}")
                failed.append(check_id)
    except KeyboardInterrupt:
        print("\nRoka gaya (Ctrl+C). Ab tak ka summary:")
    finally:
        try:
            word.Quit()
        except Exception:
            pass

    mins = (time.time() - start) / 60
    print("\n" + "=" * 40)
    print(f"Done in {mins:.1f} min")
    print(f"OK         : {len(ok)}")
    print(f"Skipped    : {len(skipped)}")
    print(f"Not found  : {len(not_found)}")
    print(f"Error      : {len(failed)}")

    problem_ids = not_found + failed
    if problem_ids:
        with open(FAILED_FILE, "w", encoding="utf-8") as f:
            f.write("\n".join(problem_ids))
        print(f"\nNot found / error wali IDs '{FAILED_FILE}' mein save ho gayi:")
        print(", ".join(problem_ids))


if __name__ == "__main__":
    main()
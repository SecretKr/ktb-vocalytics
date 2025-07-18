import csv
import re
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import os.path
from google.auth.transport.requests import Request

# === Config ===
SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive.file']
TRANSCRIPT_FILE = 'transcript/transcript.csv'
KEYWORDS_FILE = 'keywords/personal_loan.csv'
DOC_TITLE = 'Transcript with Highlights'

# === Step 1: Google Auth ===
def authenticate_google_docs():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials, prompt login
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

# === Step 2: Load and flatten transcript ===
def load_transcript(file_path):
    with open(file_path, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        chunks = [row['text'] for row in reader]
        full_text = ' '.join(chunks)
    return ' ' + full_text

# === Step 3: Load keyword sequences ===
def load_keyword_patterns(file_path):
    pattern_color_pairs = []
    with open(file_path, newline='', encoding='utf-8') as f:
        reader = csv.reader(f)
        header = next(reader)  # skip header
        for row in reader:
            color_hex = row[12].strip()  # assuming column 2 is "color"
            keywords = [w.strip() for w in row[3:7] if w.strip()]
            if not keywords:
                continue
            pattern = r'.*?'.join(map(re.escape, keywords))
            try:
                rgb = hex_to_rgb(color_hex)
            except ValueError as e:
                print(f"⚠️ Skipping invalid color: {color_hex} ({e})")
                continue
            pattern_color_pairs.append((re.compile(pattern, re.IGNORECASE | re.DOTALL), rgb))
    return pattern_color_pairs

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    if len(hex_color) != 6:
        raise ValueError(f"Invalid hex color: {hex_color}")
    r = int(hex_color[0:2], 16) / 255.0
    g = int(hex_color[2:4], 16) / 255.0
    b = int(hex_color[4:6], 16) / 255.0
    return {"red": r, "green": g, "blue": b}

# === Step 4: Create Google Doc with highlights ===
def create_doc_and_highlight(creds, full_text, pattern_color_pairs):
    docs_service = build('docs', 'v1', credentials=creds)

    # Create the document
    doc = docs_service.documents().create(body={'title': DOC_TITLE}).execute()
    doc_id = doc['documentId']

    # Insert full text
    requests = [{"insertText": {"location": {"index": 1}, "text": full_text}}]
    docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()

    # Highlight matches
    highlight_requests = []
    for pattern, rgb_color in pattern_color_pairs:
        for match in pattern.finditer(full_text):
            start = match.start()
            end = match.end()
            if start == 0 or start == end:
                print(f"⚠️ Skipping match at {start}-{end}: '{match.group()}'")
                continue

            print(f"✅ Match: '{match.group()}' [{start}–{end}], Color: {rgb_color}")

            highlight_requests.append({
                "updateTextStyle": {
                    "range": {
                        "startIndex": start + 1,
                        "endIndex": end + 1
                    },
                    "textStyle": {
                        "backgroundColor": {
                            "color": {
                                "rgbColor": rgb_color
                            }
                        }
                    },
                    "fields": "backgroundColor"
                }
            })

    if highlight_requests:
        docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': highlight_requests}).execute()

    print(f'Document created: https://docs.google.com/document/d/{doc_id}/edit')

# === Run ===
if __name__ == '__main__':
    creds = authenticate_google_docs()
    full_text = load_transcript(TRANSCRIPT_FILE)
    patterns = load_keyword_patterns(KEYWORDS_FILE)
    # print(patterns)
    create_doc_and_highlight(creds, full_text, patterns)

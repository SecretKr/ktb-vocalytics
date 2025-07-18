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
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
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
    keyword_groups = {}
    with open(file_path, newline='', encoding='utf-8') as f:
        reader = csv.reader(f)
        header = next(reader)
        for row in reader:
            if len(row) < 13:
                continue
            group_name = row[1].strip()
            color_hex = row[12].strip()
            keywords = [w.strip() for w in row[3:7] if w.strip()]

            if not group_name or not keywords:
                continue

            if group_name not in keyword_groups:
                keyword_groups[group_name] = {
                    "patterns": [],
                    "found_words": []
                }
            
            keyword_string = " ".join(keywords)
            pattern = r'.*?'.join(map(re.escape, keywords))
            try:
                rgb = hex_to_rgb(color_hex)
            except ValueError as e:
                print(f"⚠️ Skipping invalid color for group '{group_name}': {color_hex} ({e})")
                continue
            
            keyword_groups[group_name]["patterns"].append((keyword_string, re.compile(pattern, re.IGNORECASE | re.DOTALL), rgb))
            
    return keyword_groups

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    if len(hex_color) != 6:
        raise ValueError(f"Invalid hex color: {hex_color}")
    r = int(hex_color[0:2], 16) / 255.0
    g = int(hex_color[2:4], 16) / 255.0
    b = int(hex_color[4:6], 16) / 255.0
    return {"red": r, "green": g, "blue": b}

# === Step 4: Create Google Doc with highlights ===
def create_doc_and_highlight(creds, full_text, keyword_groups):
    docs_service = build('docs', 'v1', credentials=creds)

    doc = docs_service.documents().create(body={'title': DOC_TITLE}).execute()
    doc_id = doc['documentId']

    requests = [{"insertText": {"location": {"index": 1}, "text": full_text}}]
    docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()

    highlight_requests = []
    for group_name, data in keyword_groups.items():
        for keyword_string, pattern, rgb_color in data["patterns"]:
            for match in pattern.finditer(full_text):
                start = match.start()
                end = match.end()
                if start == 0 or start == end:
                    continue

                print(f"✅ Match for group '{group_name}': '{match.group()}'")
                if (keyword_string, rgb_color) not in data["found_words"]:
                    data["found_words"].append((keyword_string, rgb_color))

                highlight_requests.append({
                    "updateTextStyle": {
                        "range": {"startIndex": start + 1, "endIndex": end + 1},
                        "textStyle": {"backgroundColor": {"color": {"rgbColor": rgb_color}}},
                        "fields": "backgroundColor"
                    }
                })

    if highlight_requests:
        docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': highlight_requests}).execute()

    insert_summary_table(docs_service, doc_id, keyword_groups, full_text)

    print(f'Document created: https://docs.google.com/document/d/{doc_id}/edit')
    return keyword_groups

def insert_summary_table(docs_service, doc_id, summary_data, full_text):
    requests = [
        {'insertText': {'location': {'index': len(full_text) + 1}, 'text': '\n\nMatch Summary\n'}},
    ]
    docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()

    doc = docs_service.documents().get(documentId=doc_id).execute()
    end_of_doc_index = doc['body']['content'][-1]['endIndex']

    rows = len(summary_data) + 1
    cols = 3
    docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': [{
        'insertTable': {'location': {'index': end_of_doc_index - 1}, 'rows': rows, 'columns': cols}
    }]}).execute()

    doc = docs_service.documents().get(documentId=doc_id, fields='body(content(table))').execute()
    table_element = next((el for el in reversed(doc['body']['content']) if 'table' in el), None)
    if not table_element:
        print("Error: Could not find the inserted summary table.")
        return
    table = table_element['table']

    insert_text_requests = []
    headers = ["Sales Approach", "Found Words", "Status"]
    for c, header_text in enumerate(headers):
        cell = table['tableRows'][0]['tableCells'][c]
        cell_start_index = cell['content'][0]['startIndex']
        insert_text_requests.append({'insertText': {'location': {'index': cell_start_index}, 'text': header_text}})

    for r, (group, data) in enumerate(summary_data.items(), 1):
        cell = table['tableRows'][r]['tableCells'][0]
        cell_start_index = cell['content'][0]['startIndex']
        insert_text_requests.append({'insertText': {'location': {'index': cell_start_index}, 'text': group}})

        found_words_text = ", ".join([word for word, color in data["found_words"]])
        if found_words_text:
            cell = table['tableRows'][r]['tableCells'][1]
            cell_start_index = cell['content'][0]['startIndex']
            insert_text_requests.append({'insertText': {'location': {'index': cell_start_index}, 'text': found_words_text}})

        cell = table['tableRows'][r]['tableCells'][2]
        cell_start_index = cell['content'][0]['startIndex']
        status = "Found" if data["found_words"] else "Not Found"
        insert_text_requests.append({'insertText': {'location': {'index': cell_start_index}, 'text': status}})

    if insert_text_requests:
        docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': list(reversed(insert_text_requests))}).execute()

    # Apply formatting
    doc = docs_service.documents().get(documentId=doc_id, fields='body(content(table))').execute()
    table_element = next((el for el in reversed(doc['body']['content']) if 'table' in el), None)
    table = table_element['table']

    style_requests = []
    for r, (group, data) in enumerate(summary_data.items(), 1):
        if data["found_words"]:
            cell = table['tableRows'][r]['tableCells'][1]
            cell_start_index = cell['content'][0]['startIndex']
            
            current_pos = 0
            for i, (word, color) in enumerate(data["found_words"]):
                start_index = cell_start_index + current_pos
                end_index = start_index + len(word)
                style_requests.append({
                    "updateTextStyle": {
                        "range": {"startIndex": start_index, "endIndex": end_index},
                        "textStyle": {"backgroundColor": {"color": {"rgbColor": color}}},
                        "fields": "backgroundColor"
                    }
                })
                current_pos += len(word) + 2 # +2 for the ", "

    if style_requests:
        docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': style_requests}).execute()

def print_summary_table(summary_data):
    print("\n=== Match Summary by Group ===")
    if not summary_data:
        print("No keyword groups found.")
        return

    print(f"{'Sales Approach':<30} | {'Found Words':<40} | {'Status'}")
    print("-" * 80)
    for group, data in summary_data.items():
        status = "Found" if data["found_words"] else "Not Found"
        found_words_str = ", ".join([word for word, color in data["found_words"]])
        print(f"{group:<30} | {found_words_str:<40} | {status}")
    print("-" * 80)

# === Run ===
if __name__ == '__main__':
    creds = authenticate_google_docs()
    full_text = load_transcript(TRANSCRIPT_FILE)
    keyword_groups = load_keyword_patterns(KEYWORDS_FILE)
    summary_result = create_doc_and_highlight(creds, full_text, keyword_groups)
    print_summary_table(summary_result)

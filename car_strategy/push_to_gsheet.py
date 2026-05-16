import os, pandas as pd, json, gspread, numpy as np
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build

# Correct path to your Excel report
xls = pd.ExcelFile("car_strategy/sma_filter_report.xlsx")

# Authorize
creds_dict = json.loads(os.environ['GCP_CREDENTIALS'])
scope = ["https://www.googleapis.com/auth/spreadsheets"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)
service = build('sheets', 'v4', credentials=creds)

spreadsheet_id = "1bkwZz3tuclh_Iux_l56HpwEiYbekCNLug_4U2h_xjOM"
spreadsheet = client.open_by_key(spreadsheet_id)

for sheet_name in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet_name)

    # Clean values for JSON compliance
    df = df.replace([np.inf, -np.inf], "")
    df = df.fillna("")
    df = df.astype(str)

    try:
        ws = spreadsheet.worksheet(sheet_name)
    except:
        ws = spreadsheet.add_worksheet(title=sheet_name, rows="1000", cols="20")
    ws.clear()
    ws.update([list(df.columns)] + df.values.tolist())

    # Build formatting requests
    requests = []

    # Header row formatting
    requests.append({
        "repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1},
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": {"red": 0.12, "green": 0.30, "blue": 0.79},
                    "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}}
                }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat)"
        }
    })

    # Row colouring based on SMA Signal
    if "SMA Signal" in df.columns:
        for ri, val in enumerate(df["SMA Signal"], 1):
            color = None
            if "STRONG" in val:
                color = {"red": 0.80, "green": 0.95, "blue": 0.80}  # light green
            elif "GOOD" in val:
                color = {"red": 0.90, "green": 0.98, "blue": 0.90}  # pale green
            elif "OVERBOUGHT" in val:
                color = {"red": 0.95, "green": 0.80, "blue": 0.80}  # light red
            elif "BELOW" in val:
                color = {"red": 0.98, "green": 0.90, "blue": 0.70}  # orange
            if color:
                requests.append({
                    "repeatCell": {
                        "range": {"sheetId": ws.id, "startRowIndex": ri, "endRowIndex": ri+1},
                        "cell": {"userEnteredFormat": {"backgroundColor": color}},
                        "fields": "userEnteredFormat(backgroundColor)"
                    }
                })

    if requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests}
        ).execute()

print("✅ SMA Filter workbook pushed with colour indicators to Google Sheet")


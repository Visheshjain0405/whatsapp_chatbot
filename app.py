import pandas as pd
import requests
import time
from flask import Flask, request, jsonify
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
import os

# Load environment variables from .env file
load_dotenv()

# Flask App
app = Flask(__name__)

# WhatsApp API Credentials
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN")
PHONE_NUMBER_ID = os.getenv("PHONE_NUMBER_ID")
VERIFY_TOKEN = os.getenv("VERIFY_TOKEN")

# Google Sheets Setup
SERVICE_ACCOUNT_FILE = os.getenv("SERVICE_ACCOUNT_FILE")
SHEET_NAME = os.getenv("SHEET_NAME")

# Authenticate and load sheet
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scope)
client = gspread.authorize(creds)
sheet = client.open(SHEET_NAME).sheet1
data = sheet.get_all_records()
df = pd.DataFrame(data)

# Helper to update cell in Google Sheet
def update_google_sheet_cell(row_index, col_name, value):
    col_index = df.columns.get_loc(col_name) + 1
    sheet.update_cell(row_index + 2, col_index, value)

# WhatsApp message sender
def send_whatsapp_message(to, first_name, training_title, date, time):
    date_str = date.strftime('%d %B %Y') if hasattr(date, 'strftime') else date
    time_str = time.strftime('%I:%M %p %Z') if hasattr(time, 'strftime') else time

    url = f"https://graph.facebook.com/v22.0/{PHONE_NUMBER_ID}/messages"
    headers = {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Content-Type": "application/json"
    }
    payload = {
        "messaging_product": "whatsapp",
        "to": to,
        "type": "template",
        "template": {
            "name": "bdc_training",
            "language": {"code": "en"},
            "components": [
                {
                    "type": "body",
                    "parameters": [
                        {"type": "text", "text": f"{first_name}"},
                        {"type": "text", "text": date_str},
                        {"type": "text", "text": time_str},
                        {"type": "text", "text": training_title}
                    ]
                }
            ]
        }
    }
    try:
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()
        print(f"Message sent to {to}: {response.text}")
        return True
    except requests.exceptions.RequestException as e:
        print(f"Failed to send message to {to}: {e.response.text if e.response else str(e)}")
        return False

# Confirmation message sender
def send_confirmation_message(to, first_name, training_title, date, time):
    date_str = date.strftime('%d %B %Y') if hasattr(date, 'strftime') else date
    time_str = time.strftime('%I:%M %p %Z') if hasattr(time, 'strftime') else time

    url = f"https://graph.facebook.com/v22.0/{PHONE_NUMBER_ID}/messages"
    headers = {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Content-Type": "application/json"
    }
    payload = {
        "messaging_product": "whatsapp",
        "to": to,
        "type": "template",
        "template": {
            "name": "training_confirmation",
            "language": {"code": "en_US"},
            "components": [
                {
                    "type": "body",
                    "parameters": [
                        {"type": "text", "text": f"{first_name}"},
                        {"type": "text", "text": date_str},
                        {"type": "text", "text": time_str},
                        {"type": "text", "text": training_title}
                    ]
                }
            ]
        }
    }
    try:
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()
        print(f"Confirmation sent to {to}: {response.text}")
        return True
    except requests.exceptions.RequestException as e:
        print(f"Failed to send confirmation to {to}: {e.response.text if e.response else str(e)}")
        return False

# Send initial WhatsApp messages
def send_messages():
    for index, row in df.iterrows():
        if row["WhatsApp Msg Status"] == "Not Send":
            mobile = str(row["Mobile Number"])
            print(f"Attempting to send to {mobile}")
            if send_whatsapp_message(mobile, row["FirstName"], row["Training Title"], row["Date"], row["Time"]):
                print(f"Message sent to {row['FirstName']} {row['LastName']}")
                update_google_sheet_cell(index, "WhatsApp Msg Status", "Send")
            else:
                print(f"Failed to send to {row['FirstName']} {row['LastName']}")
            time.sleep(1)

# Webhook handler
@app.route('/webhook', methods=['GET', 'POST'])
def webhook():
    if request.method == 'GET':
        verify_token = request.args.get('hub.verify_token')
        challenge = request.args.get('hub.challenge')
        if verify_token == VERIFY_TOKEN:
            return challenge, 200
        return "Verification failed", 403

    elif request.method == 'POST':
        data = request.get_json()
        if data.get('entry') and data['entry'][0].get('changes'):
            message = data['entry'][0]['changes'][0]['value'].get('messages', [{}])[0]
            if message.get('type') == 'button':
                payload = message['button'].get('payload')
                from_number = message['from']
                for index, row in df.iterrows():
                    if str(row["Mobile Number"]) == from_number:
                        if payload == "YES":
                            update_google_sheet_cell(index, "Confirmation", "Yes")
                            send_confirmation_message(from_number, row["FirstName"], row["Training Title"], row["Date"], row["Time"])
                        elif payload == "NO":
                            update_google_sheet_cell(index, "Confirmation", "No")
                        print(f"Updated {row['FirstName']} {row['LastName']} with confirmation: {payload}")
                        break
        return jsonify({"status": "success"}), 200

# Entry point
if __name__ == "__main__":
    send_messages()
    app.run(port=5000, debug=True)

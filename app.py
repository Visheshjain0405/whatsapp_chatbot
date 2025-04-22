import pandas as pd
import requests
import time
from flask import Flask, request, jsonify

app = Flask(__name__)
ACCESS_TOKEN = "EAAOj6ZAIAuUYBO2ti6mZAgPna95G4mc2YfxiNzZAnPPd5kZB0paahFmJToGdhp90rQyJfaUBZAZCJZCDrWRJKP71ZAIigCUgacCcNjO0ZBCz9MMDn0QHwlmknCDaGci5NBCqmQbffuIdmFQ35ZBn1BZBuOWVNDPnHCeADORFYZCEDW7RhlXAu9M0RnROBtPMgWc6ODNGopG1aD83di0yKIJZAiby6VcK13a8ZD"
PHONE_NUMBER_ID = "624030817461040"  # Verified working
EXCEL_FILE = "contacts.xlsx"
VERIFY_TOKEN = "BDC_Surat"  # Replace with your custom verify token (e.g., "my_secret_token")

# Load Excel data
df = pd.read_excel(EXCEL_FILE)

# Function to send WhatsApp template message
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
            "name": "bdc_training",  # Updated to use bdc_training
            "language": {"code": "en"},
            "components": [
                {
                    "type": "body",
                    "parameters": [
                        {"type": "text", "text": f"{first_name}"},
                        {"type": "text", "text": date_str},
                        {"type": "text", "text": time_str},
                        {"type": "text", "text": training_title}  # Fixed to use training_title
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

# Function to send confirmation message
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

# Send messages to all unprocessed contacts
def send_messages():
    for index, row in df.iterrows():
        if row["WhatsApp Msg Status"] == "Not Send":
            mobile = str(row["Mobile Number"])
            print(f"Attempting to send to {mobile}")
            if send_whatsapp_message(mobile, row["FirstName"], row["Training Title"], row["Date"], row["Time"]):
                print(f"Message sent to {row['FirstName']} {row['LastName']}")
                try:
                    df.at[index, "WhatsApp Msg Status"] = "Send"
                    df.to_excel(EXCEL_FILE, index=False)
                    print(f"Updated {EXCEL_FILE} for {row['FirstName']} {row['LastName']}")
                except PermissionError as e:
                    print(f"Permission denied while saving {EXCEL_FILE}: {e}. Please close the file or run with admin privileges.")
                except Exception as e:
                    print(f"Error saving {EXCEL_FILE}: {e}")
            else:
                print(f"Failed to send to {row['FirstName']} {row['LastName']}")
            time.sleep(1)  # Avoid rate limiting

# Webhook to handle responses
@app.route('/webhook', methods=['GET', 'POST'])
def webhook():
    if request.method == 'GET':
        # Verify webhook
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
                # Update Excel based on response
                for index, row in df.iterrows():
                    if str(row["Mobile Number"]) == from_number:
                        if payload == "YES":
                            df.at[index, "Confirmation"] = "Yes"
                            send_confirmation_message(from_number, row["FirstName"], row["Training Title"], row["Date"], row["Time"])
                        elif payload == "NO":
                            df.at[index, "Confirmation"] = "No"
                        try:
                            df.to_excel(EXCEL_FILE, index=False)
                            print(f"Updated {row['FirstName']} {row['LastName']} with confirmation: {payload}")
                        except PermissionError as e:
                            print(f"Permission denied while saving {EXCEL_FILE}: {e}. Please close the file or run with admin privileges.")
                        except Exception as e:
                            print(f"Error saving {EXCEL_FILE}: {e}")
                        break
        return jsonify({"status": "success"}), 200

if __name__ == "__main__":
    send_messages()  # Send initial messages
    app.run(port=5000, debug=True)  # Start webhook server with debug mode

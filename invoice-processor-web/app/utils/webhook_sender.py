import requests
from decouple import config

def send_to_webhook(data):
    webhook_url = config("MAKE_WEBHOOK_URL")
    try:
        response = requests.post(webhook_url, json=data)
        return response.status_code == 200
    except Exception as e:
        print(f"Webhook Error: {e}")
        return False
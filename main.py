import requests
import json
import logging
import schedule
import time

# Constants
CLIENT_ID = "b8ebdc42-2414-425d-95e2-d70b9f94fb43"
CLIENT_SECRET = "ynC8Q~j1e9HkIjdpJXVfzHHhCZllaVxdQ~2J8bGi"
TENANT_ID = "57bd375a-8f5a-4585-8cd4-9c82ba31f845"
GRAPH_API = "https://graph.microsoft.com/v1.0"
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

ACCESS_TOKEN = None

logging.basicConfig(level=logging.INFO)

# Get access token
def get_access_token():
    global ACCESS_TOKEN
    url = TOKEN_URL
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }
    response = requests.post(url, data=data)
    if response.status_code != 200:
        raise Exception(f"Failed to get token: {response.status_code}, {response.text}")
    ACCESS_TOKEN = response.json().get("access_token")
    logging.info("Access token obtained successfully.")

# Get folder ID
def get_folder_id(user_id, parent_folder_name, target_folder_name):
    # Step 1: Get the parent folder ID
    parent_folder_url = f"{GRAPH_API}/users/{user_id}/mailFolders"
    headers = {"Authorization": f"Bearer {ACCESS_TOKEN}"}
    response = requests.get(parent_folder_url, headers=headers)
    if response.status_code != 200:
        raise Exception(f"Failed to get mail folders: {response.status_code}, {response.text}")
    folders = response.json().get("value", [])
    
    parent_folder_id = next((folder["id"] for folder in folders if folder["displayName"] == parent_folder_name), None)
    if not parent_folder_id:
        raise Exception(f"Parent folder {parent_folder_name} not found.")

    # Step 2: Get the child folder ID
    child_folder_url = f"{GRAPH_API}/users/{user_id}/mailFolders/{parent_folder_id}/childFolders"
    response = requests.get(child_folder_url, headers=headers)
    if response.status_code != 200:
        raise Exception(f"Failed to get child folders: {response.status_code}, {response.text}")
    child_folders = response.json().get("value", [])

    target_folder_id = next((folder["id"] for folder in child_folders if folder["displayName"] == target_folder_name), None)
    if not target_folder_id:
        raise Exception(f"Folder {target_folder_name} not found in {parent_folder_name}.")
    
    return target_folder_id

# Flag email
def flag_email(user_id, message_id):
    url = f"{GRAPH_API}/users/{user_id}/messages/{message_id}"
    headers = {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Content-Type": "application/json"
    }
    data = {
        "flag": {
            "flagStatus": "flagged"
        }
    }
    response = requests.patch(url, headers=headers, json=data)
    if response.status_code != 200:
        raise Exception(f"Failed to flag email: {response.status_code}, {response.text}")
    logging.info(f"Email with ID {message_id} flagged successfully.")

# Monitor folder
def monitor_folder(user_id, folder_id):
    url = f"{GRAPH_API}/users/{user_id}/mailFolders/{folder_id}/messages?$filter=flag/flagStatus eq 'notFlagged'"
    headers = {"Authorization": f"Bearer {ACCESS_TOKEN}"}
    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        raise Exception(f"Failed to get messages: {response.status_code}, {response.text}")
    messages = response.json().get("value", [])
    
    for message in messages:
        # Forward the email
        forward_email(user_id, message["id"])

        # Flag the email
        flag_email(user_id, message["id"])

# Forward email
def forward_email(user_id, message_id):
    url = f"{GRAPH_API}/users/{user_id}/messages/{message_id}/forward"
    headers = {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Content-Type": "application/json"
    }
    data = {
        "comment": "Forwarding this email as requested.",
        "toRecipients": [
            {"emailAddress": {"address": "remotesupportnederlandez31400.4170@to-zenvoices.com"}},
            {"emailAddress": {"address": "5a75313s@inkoop.exactonline.nl"}}
        ]
    }
    response = requests.post(url, headers=headers, json=data)
    if response.status_code != 202:
        raise Exception(f"Failed to forward email: {response.status_code}, {response.text}")
    logging.info(f"Email with ID {message_id} forwarded successfully.")


# Main function with schedule
def main():
    user_id = "finance@remotesupportnederland.nl"
    parent_folder_name = "Inbox"
    target_folder_name = "Facturen betaald"
    
    def job():
        try:
            get_access_token()
            folder_id = get_folder_id(user_id, parent_folder_name, target_folder_name)
            monitor_folder(user_id, folder_id)
            logging.info("Emails processed successfully.")
        except Exception as e:
            logging.error(f"Error: {e}")
    job()
    # Schedule the job to run every 20 seconds
    schedule.every(20).seconds.do(job)
    
    logging.info("Starting scheduled tasks...")
    while True:
        schedule.run_pending()
        time.sleep(5)

if __name__ == "__main__":
    main()

import requests
import os

# Insert your valid access token here
ACCESS_TOKEN = "YOUR_ACCESS_TOKEN"

# Number of emails to retrieve
MAX_EMAILS = 10

# Microsoft Graph API endpoints for received and sent emails
INBOX_ENDPOINT = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages"
SENT_ENDPOINT = "https://graph.microsoft.com/v1.0/me/mailFolders/sentitems/messages"

# Set the Authorization header
headers = {
    "Authorization": f"Bearer {ACCESS_TOKEN}",
    "Content-Type": "application/json"
}

# Parameters for fetching emails
params = {
    "$top": MAX_EMAILS,
    "$orderby": "receivedDateTime DESC",
    "$select": "subject,from,bodyPreview,body,attachments"
}

# Directory to save downloaded attachments
DOWNLOAD_DIR = "attachments"

# Function to search within email body and subject
def search_in_email(email, keyword):
    subject = email.get('subject', '').lower()
    body = email.get('bodyPreview', '').lower()  # Quick preview of body

    if keyword.lower() in subject or keyword.lower() in body:
        return True
    return False

# Function to download attachments from an email
def download_attachments(attachments, email_subject):
    if not os.path.exists(DOWNLOAD_DIR):
        os.makedirs(DOWNLOAD_DIR)

    for attachment in attachments:
        # Check if the attachment is a file and not inline content (images, etc.)
        if attachment['@odata.type'] == "#microsoft.graph.fileAttachment":
            attachment_name = attachment['name']
            attachment_content = attachment['contentBytes']

            # Save to disk
            file_path = os.path.join(DOWNLOAD_DIR, attachment_name)
            with open(file_path, 'wb') as f:
                f.write(attachment_content.decode('base64'))
            print(f"[+] Downloaded attachment: {attachment_name}")

# Function to fetch emails from a given endpoint (inbox or sent)
def fetch_emails(endpoint, label, search_keyword=None):
    print(f"[+] Fetching last {MAX_EMAILS} {label} emails...")

    response = requests.get(endpoint, headers=headers, params=params)

    if response.status_code == 200:
        emails = response.json().get('value', [])
        print(f"[+] Retrieved {len(emails)} {label} emails.\n")

        for idx, email in enumerate(emails, 1):
            subject = email.get('subject', '(No Subject)')
            sender = email.get('from', {}).get('emailAddress', {}).get('address', '(Unknown Sender)')
            body_preview = email.get('bodyPreview', '(No Body Preview)')
            body = email.get('body', {}).get('content', '')
            attachments = email.get('attachments', [])

            # If a search keyword is provided, search within subject and body
            if search_keyword and not search_in_email(email, search_keyword):
                continue

            print(f"--- {label.capitalize()} Email #{idx} ---")
            print(f"From   : {sender}")
            print(f"Subject: {subject}")
            print(f"Snippet: {body_preview}")
            print(f"Body   : {body[:200]}...")  # Print only first 200 chars of the body
            print(f"Attachments: {len(attachments)}")
            print("----------------------\n")

            # Download attachments if available
            if attachments:
                download_attachments(attachments, subject)
    else:
        print(f"[-] Failed to retrieve {label} emails. Status Code: {response.status_code}")
        print(response.text)

if __name__ == "__main__":
    search_keyword = input("Enter a keyword to search for in subject and body (or leave blank to skip): ")
    
    # Fetch and display emails
    fetch_emails(INBOX_ENDPOINT, "received", search_keyword)
    fetch_emails(SENT_ENDPOINT, "sent", search_keyword)

import pandas as pd
import requests
import msal
import time
import random
import re
import imaplib
import email
from datetime import datetime

# ==========================================
# 1. CONFIGURATION
# ==========================================
SENDER_EMAIL  = "admin@mobifirst.co"

# ✅ Paste your Azure App values here:
CLIENT_ID     = "dabde45a-de62-44f1-a15d-1572327a9302"
CLIENT_SECRET = "a5ca72d8-f1af-4c3a-ac8a-5727d35fe0bb"
TENANT_ID     = "f14f07b0-a186-41e6-a3b4-19cfd15af98c"

DATABASE_FILE        = "MobiFirst_Master_List.xlsx"
BATCH_SIZE           = 500
BATCH_DELAY_SECONDS  = 1800   # 30 mins between batches
EMAIL_DELAY_SECONDS  = 2      # 2 sec between emails
DAILY_LIMIT          = 10000
PAUSE_24H_SECONDS    = 86400

UNSUBSCRIBE_URL = "https://mobifirst.co/unsubscribe"

# ==========================================
# 2. EMAIL TEMPLATES
# ==========================================
def build_email_body(body_content: str) -> str:
    return f"""
    {body_content}
    <br><br>
    <hr style="border:none;border-top:1px solid #eee;margin:20px 0;">
    <p style="font-size:11px;color:#999;text-align:center;">
        You're receiving this because you previously used MobiFirst.<br>
        <a href="{UNSUBSCRIBE_URL}" style="color:#999;">Unsubscribe</a> | 
        MobiFirst
    </p>
    """

templates = [
    {
        "Subject": "You're invited back",
        "Body": "Hey there,<br><br>A lot has changed and MobiFirst just leveled up and we want you back before the Al wave gets even bigger.<br><br>On June 19th at 6:30 PM, we're hosting a special webinar to reveal the new MobiFirst platform overhaul.<br><br>Register here:<br>https://access.encorewebinars.com/webinar/registration/6a1868fb2ca928add624613d<br><br>We'd love to see you again."
    },
    {
        "Subject": "You won't believe what we rebuilt",
        "Body": "Hey,<br><br>You once used MobiFirst - but what we're about to show you is on a completely different level.<br><br>Join us June 19th at 6:30 PM for a full walkthrough of the new AI site builder and our new pricing model.<br><br>Save your seat:<br>https://access.encorewebinars.com/webinar/registration/6a1868fb2ca928add624613d"
    },
    {
        "Subject": "The time to come back is right now",
        "Body": "Hey,<br><br>Al is moving fast and the people who win are the ones who move with it. On June 19th at 6:30 PM, we're hosting a special MobiFirst comeback webinar.<br><br>If you've been waiting for the right moment to return... this is it.<br><br>Register now and see what's new:<br>https://access.encorewebinars.com/webinar/registration/6a1868fb2ca928add624613d"
    },
    {
        "Subject": "A completely new MobiFirst is waiting for you",
        "Body": "Hey there,<br><br>We've spent months rebuilding MobiFirst from the ground up and we want you to see it.<br><br>Join us June 19th at 6:30 PM for a full reveal of the new AI Site Builder and Reseller Plans at User Pricing.<br><br>Reserve your spot:<br>https://access.encorewebinars.com/webinar/registration/6a1868fb2ca928add624613d"
    },
    {
        "Subject": "We'd love to have you back",
        "Body": "Hey,<br><br>We're hosting a special MobiFirst comeback webinar on June 19th at 6:30 PM, and you're invited.<br><br>The Al boom is happening right now and we want you to be part of it.<br><br>Register here:<br>https://access.encorewebinars.com/webinar/registration/6a1868fb2ca928add624613d"
    }
]

# ==========================================
# 3. GRAPH API TOKEN
# ==========================================
_token_cache = {"token": None, "expires_at": 0}

def get_access_token() -> str:
    """Gets a cached token or fetches a new one."""
    if _token_cache["token"] and time.time() < _token_cache["expires_at"] - 60:
        return _token_cache["token"]

    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise Exception(f"❌ Token error: {result.get('error_description')}")

    _token_cache["token"] = result["access_token"]
    _token_cache["expires_at"] = time.time() + result.get("expires_in", 3600)
    return _token_cache["token"]

# ==========================================
# 4. SEND EMAIL VIA GRAPH API
# ==========================================
def send_email_graph(to: str, subject: str, body: str) -> bool:
    """Sends a single email via Microsoft Graph API (no SMTP needed)."""
    try:
        token = get_access_token()
        response = requests.post(
            f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json"
            },
            json={
                "message": {
                    "subject": subject,
                    "body": {
                        "contentType": "HTML",
                        "content": body
                    },
                    "toRecipients": [
                        {"emailAddress": {"address": to}}
                    ]
                },
                "saveToSentItems": "true"
            },
            timeout=30
        )
        return response.status_code == 202
    except Exception as e:
        print(f"    Graph API error: {e}")
        return False

# ==========================================
# 5. BOUNCE HANDLER (IMAP)
# ==========================================
def process_bounces(df):
    print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Checking inbox for bounces...")
    try:
        mail = imaplib.IMAP4_SSL("outlook.office365.com")
        mail.login(SENDER_EMAIL, "your-password-here")  # IMAP still needs password
        mail.select("inbox")

        status, messages = mail.search(None, '(SUBJECT "Undeliverable")')
        email_ids = messages[0].split()

        if not email_ids:
            print("  No new bounces found.")
            mail.logout()
            return df

        print(f"  Found {len(email_ids)} bounce reports. Updating database...")
        email_regex = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'

        for e_id in email_ids:
            _, msg_data = mail.fetch(e_id, '(RFC822)')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    body = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            if part.get_content_type() == "text/plain":
                                body += str(part.get_payload(decode=True))
                    else:
                        body = str(msg.get_payload(decode=True))

                    found_emails = re.findall(email_regex, body)
                    for failed_email in found_emails:
                        failed_email = failed_email.lower().strip()
                        if failed_email in df['Email Address'].str.lower().values:
                            df.loc[df['Email Address'].str.lower() == failed_email, 'Status'] = 'Bounced'

            mail.store(e_id, '+FLAGS', '\\Deleted')

        mail.expunge()
        mail.logout()
        print("  Bounce processing complete.")
    except Exception as e:
        print(f"  Bounce check skipped: {e}")

    return df

# ==========================================
# 6. MAIN ENGINE
# ==========================================
def main_engine():
    # Validate Excel columns
    try:
        df_check = pd.read_excel(DATABASE_FILE)
        required_columns = ['Email Address', 'Status']
        missing = [col for col in required_columns if col not in df_check.columns]
        if missing:
            print(f"❌ Missing columns in Excel: {missing}")
            return
    except Exception as e:
        print(f"❌ Cannot open database: {e}")
        return

    # Validate Graph API credentials
    if "your-client-id-here" in CLIENT_ID:
        print("❌ Please update CLIENT_ID, CLIENT_SECRET, and TENANT_ID in the script!")
        return

    print("=" * 55)
    print("  MobiFirst AI Sender Engine — Starting Up")
    print("=" * 55)
    print(f"  Sender      : {SENDER_EMAIL}")
    print(f"  Mode        : Microsoft Graph API ✅")
    print(f"  Database    : {DATABASE_FILE}")
    print(f"  Batch Size  : {BATCH_SIZE} emails")
    print(f"  Email Delay : {EMAIL_DELAY_SECONDS} seconds")
    print(f"  Batch Rest  : {BATCH_DELAY_SECONDS/60:.0f} minutes")
    print(f"  Daily Limit : {DAILY_LIMIT}")
    print("=" * 55)

    # Test token on startup
    try:
        get_access_token()
        print("  ✅ Graph API token obtained successfully!")
    except Exception as e:
        print(f"  ❌ Graph API auth failed: {e}")
        print("  Check your CLIENT_ID, CLIENT_SECRET, TENANT_ID values.")
        return

    total_sent_today = 0

    while True:
        try:
            df = pd.read_excel(DATABASE_FILE)
        except Exception as e:
            print(f"❌ Failed to load database: {e}")
            break

        df = process_bounces(df)

        pending_indices = df.index[df['Status'] == 'Pending'].tolist()
        if not pending_indices:
            # Final summary
            sent_count    = len(df[df['Status'] == 'Sent'])
            bounced_count = len(df[df['Status'] == 'Bounced'])
            failed_count  = len(df[df['Status'] == 'Failed'])
            invalid_count = len(df[df['Status'] == 'Invalid'])
            total_count   = len(df)

            print(f"\n{'='*55}")
            print(f"  🎉 CAMPAIGN COMPLETE — {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"{'='*55}")
            print(f"  ✅ Total Sent    : {sent_count}")
            print(f"  ❌ Total Bounced : {bounced_count}")
            print(f"  ⚠️  Total Failed  : {failed_count}")
            print(f"  🚫 Total Invalid : {invalid_count}")
            print(f"  📋 Total Records : {total_count}")
            print(f"  📈 Success Rate  : {round((sent_count/total_count)*100, 1)}%")
            print(f"{'='*55}")
            df.to_excel(DATABASE_FILE, index=False)
            break

        if total_sent_today >= DAILY_LIMIT:
            print(f"\n⚠️  Daily limit of {DAILY_LIMIT} reached.")
            print(f"   Sleeping 24 hours. Keep this window open...")
            df.to_excel(DATABASE_FILE, index=False)
            time.sleep(PAUSE_24H_SECONDS)
            total_sent_today = 0
            continue

        batch_indices = pending_indices[:BATCH_SIZE]
        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Starting batch — {len(batch_indices)} emails...")
        print(f"  Progress today: {total_sent_today}/{DAILY_LIMIT}")

        batch_success = 0
        for idx in batch_indices:
            if total_sent_today >= DAILY_LIMIT:
                break

            recipient_email = str(df.at[idx, 'Email Address']).strip()

            if '@' not in recipient_email or '.' not in recipient_email:
                df.at[idx, 'Status'] = 'Invalid'
                continue

            template = random.choice(templates)
            full_body = build_email_body(template['Body'])

            success = send_email_graph(recipient_email, template['Subject'], full_body)

            if success:
                df.at[idx, 'Status'] = 'Sent'
                batch_success += 1
                total_sent_today += 1
                print(f"  ✅ [{total_sent_today}/{DAILY_LIMIT}] Sent → {recipient_email}")
            else:
                df.at[idx, 'Status'] = 'Failed'
                print(f"  ❌ Failed → {recipient_email}")

            time.sleep(EMAIL_DELAY_SECONDS)

        df.to_excel(DATABASE_FILE, index=False)

        # Batch summary
        sent_count    = len(df[df['Status'] == 'Sent'])
        pending_count = len(df[df['Status'] == 'Pending'])
        bounced_count = len(df[df['Status'] == 'Bounced'])
        failed_count  = len(df[df['Status'] == 'Failed'])
        invalid_count = len(df[df['Status'] == 'Invalid'])
        total_count   = len(df)

        print(f"\n{'='*55}")
        print(f"  📊 BATCH SUMMARY — {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"{'='*55}")
        print(f"  ✅ Sent (this batch) : {batch_success}")
        print(f"  ✅ Total Sent        : {sent_count} / {total_count}")
        print(f"  ⏳ Pending           : {pending_count}")
        print(f"  ❌ Bounced           : {bounced_count}")
        print(f"  ⚠️  Failed            : {failed_count}")
        print(f"  🚫 Invalid           : {invalid_count}")
        print(f"  📈 Progress          : {round((sent_count/total_count)*100, 1)}%")
        print(f"  📬 Sent today        : {total_sent_today} / {DAILY_LIMIT}")
        print(f"{'='*55}")
        print(f"  💤 Resting for {BATCH_DELAY_SECONDS/60:.0f} minutes before next batch...")
        print(f"{'='*55}\n")

        time.sleep(BATCH_DELAY_SECONDS)


if __name__ == "__main__":
    main_engine()

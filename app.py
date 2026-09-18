"""
Outlook spam-filter bot using Microsoft Graph change notifications.

Flow:
  1. Authenticate once via device code (MSAL), token cached to disk.
  2. Create a Graph subscription on the Inbox for "created" events.
  3. Run a small Flask server that:
       - answers the validation handshake Graph sends when creating/renewing
         the subscription
       - receives notifications, fetches the new message, runs spam logic,
         and moves/deletes it
  4. A background thread renews the subscription before it expires
     (max lifetime for messages is ~4230 minutes / ~2.9 days).

Everything here is free: Graph API calls, MSAL, Flask. You still need a
public HTTPS URL pointing at this server (see cloudflared/ngrok notes).
"""

import json
import os
import threading
import time
from datetime import datetime, timedelta, timezone

import msal
import requests
from flask import Flask, request, Response
from waitress import serve

# ---------------------------------------------------------------------------
# Configuration - set these via environment variables (see .env.example)
# ---------------------------------------------------------------------------

CLIENT_ID = os.environ["CLIENT_ID"]
TENANT_ID = os.environ.get("TENANT_ID", "common")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Mail.ReadWrite", "Mail.Read", "User.Read"]

# The public HTTPS URL that forwards to this script. Two ways to set it:
#  - NOTIFICATION_URL env var, set directly (use this for a named/stable tunnel)
#  - TUNNEL_URL_FILE, written by the cloudflared container with the current
#    quick-tunnel URL (used automatically if NOTIFICATION_URL isn't set)
NOTIFICATION_URL = os.environ.get("NOTIFICATION_URL")
TUNNEL_URL_FILE = os.environ.get("TUNNEL_URL_FILE", "/shared/tunnel_url.txt")

# A secret string Graph will echo back in each notification so you can
# verify it's really Graph calling you. Make this random and keep it secret.
CLIENT_STATE = os.environ["CLIENT_STATE"]

# All state lives under /data, which is a mounted volume so it survives
# container restarts and can move to a new machine (e.g. the Raspberry Pi).
DATA_DIR = os.environ.get("DATA_DIR", "/data")
TOKEN_CACHE_FILE = os.path.join(DATA_DIR, "token_cache.bin")
SUBSCRIPTION_FILE = os.path.join(DATA_DIR, "subscription.json")
GRAPH_ROOT = "https://graph.microsoft.com/v1.0"

# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------


def load_token_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        cache.deserialize(open(TOKEN_CACHE_FILE, "r").read())
    return cache


def save_token_cache(cache):
    if cache.has_state_changed:
        with open(TOKEN_CACHE_FILE, "w") as f:
            f.write(cache.serialize())


def get_access_token():
    cache = load_token_cache()
    app = msal.PublicClientApplication(
        CLIENT_ID, authority=AUTHORITY, token_cache=cache)

    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        # First run: interactive device code login
        flow = app.initiate_device_flow(scopes=SCOPES)
        print(get_time(), "FLOW:", flow)
        if "user_code" not in flow:
            raise RuntimeError(f"Failed to create device flow: {flow}")
        # e.g. "Go to https://microsoft.com/devicelogin and enter code XXXX"
        print(get_time(), flow["message"])
        result = app.acquire_token_by_device_flow(flow)

    save_token_cache(cache)

    if "access_token" not in result:
        raise RuntimeError(f"Could not get token: {result}")
    return result["access_token"]


def graph_headers():
    return {
        "Authorization": f"Bearer {get_access_token()}",
        "Content-Type": "application/json",
    }

# ---------------------------------------------------------------------------
# Subscription management
# ---------------------------------------------------------------------------


def wait_for_tunnel_url(path, timeout=120):
    """Poll the file cloudflared writes its assigned URL to, until it appears."""
    waited = 0
    last_content = None
    while waited < timeout:
        if os.path.exists(path):
            content = open(path).read().strip()
            if content:
                if content != last_content:
                    print(get_time(), f"Tunnel URL detected: {content}")
                    last_content = content
                return content
        time.sleep(2)
        waited += 2
    raise RuntimeError(
        f"Timed out after {timeout}s waiting for {path} - is the cloudflared container running?"
    )


def wait_for_tunnel_reachable(base_url, timeout=300):
    """Quick tunnel hostnames can take a few seconds to become resolvable
    worldwide after cloudflared prints them. Poll from inside our own
    container until the URL actually responds, so we don't hand Graph a
    dead hostname (which it'll reject with a DNS resolution error)."""
    waited = 0
    print(get_time(), "wait_for_tunnel_reachable:", base_url)
    while waited < timeout:
        print(get_time(), f"waited {waited} seconds")
        try:
            # Any response (even 404/405) proves DNS + routing are working.
            requests.get(base_url, timeout=5)
            print(get_time(), f"Tunnel confirmed reachable: {base_url}")
            return
        except requests.exceptions.RequestException:
            pass
        time.sleep(5)
        waited += 5
    raise RuntimeError(
        f"Tunnel at {base_url} never became reachable after {timeout}s")


def delete_subscription(sub_id):
    resp = requests.delete(
        f"{GRAPH_ROOT}/subscriptions/{sub_id}", headers=graph_headers())
    if resp.status_code not in (204, 404):
        print(get_time(),
              f"Warning: could not delete old subscription {sub_id}: {resp.status_code} {resp.text}")


def create_subscription_junkemail():
    expiration = (datetime.now(timezone.utc) +
                  timedelta(minutes=4200)).isoformat()
    body = {
        "changeType": "created",
        "notificationUrl": NOTIFICATION_URL,
        "resource": "me/mailFolders('junkemail')/messages",
        "expirationDateTime": expiration,
        "clientState": CLIENT_STATE,
    }
    resp = requests.post(f"{GRAPH_ROOT}/subscriptions",
                         headers=graph_headers(), json=body)
    if not resp.ok:
        print(get_time(),
              f"Subscription creation failed: {resp.status_code} {resp.text}")
    resp.raise_for_status()
    sub = resp.json()
    with open(SUBSCRIPTION_FILE, "w") as f:
        json.dump(sub, f)
    print(get_time(),
          f"Subscription created, id={sub['id']}, expires {sub['expirationDateTime']}")
    return sub


def renew_subscription(sub_id):
    expiration = (datetime.now(timezone.utc) +
                  timedelta(minutes=4200)).isoformat()
    resp = requests.patch(
        f"{GRAPH_ROOT}/subscriptions/{sub_id}",
        headers=graph_headers(),
        json={"expirationDateTime": expiration},
    )
    resp.raise_for_status()
    print(get_time(), f"Subscription {sub_id} renewed until {expiration}")


def subscription_renewal_loop():
    while True:
        time.sleep(60 * 60 * 24)  # check once a day
        try:
            if os.path.exists(SUBSCRIPTION_FILE):
                sub = json.load(open(SUBSCRIPTION_FILE))
                renew_subscription(sub["id"])
        except Exception as e:
            print(get_time(), f"Renewal failed, recreating subscription: {e}")
            try:
                create_subscription_junkemail()
            except Exception as e2:
                print(get_time(), f"Recreate also failed: {e2}")

# ---------------------------------------------------------------------------
# Spam logic - customize this function
# ---------------------------------------------------------------------------


SPAM_KEYWORDS = ["free money", "act now",
                 "wire transfer", "you have won", "crypto giveaway"]
# e.g. ["mycompany.com"] - never flag these as spam
TRUSTED_DOMAINS = ["microsoft.com"]


def get_header(message, header_name):
    headers = message.get("internetMessageHeaders", []) or []
    for h in headers:
        if h["name"].lower() == header_name.lower():
            return h["value"]
    return None


def is_spam(message):
    sender = (message.get("from", {}).get(
        "emailAddress", {}).get("address") or "").lower()
    subject = (message.get("subject") or "").lower()
    body_preview = (message.get("bodyPreview") or "").lower()

    if subject.endswith("test_spam"):
        return True

    if any(sender.endswith("@" + d) for d in TRUSTED_DOMAINS):
        return False
    if not sender.endswith(".com"):
        return True

    # Use Microsoft's own spam confidence level (SCL) header if present.
    # SCL ranges roughly -1 (safe) to 9 (very likely spam); >= 5 is usually junk.
    scl = get_header(message, "X-MS-Exchange-Organization-SCL")
    if scl is not None:
        try:
            if int(scl) >= 5:
                return True
        except ValueError:
            pass

    text = f"{subject} {body_preview}"
    if any(keyword in text for keyword in SPAM_KEYWORDS):
        return True

    return False

# ---------------------------------------------------------------------------
# Message actions
# ---------------------------------------------------------------------------


def fetch_message(message_id):
    resp = requests.get(
        f"{GRAPH_ROOT}/me/messages/{message_id}"
        "?$select=subject,bodyPreview,from,internetMessageHeaders",
        headers=graph_headers(),
    )
    resp.raise_for_status()
    return resp.json()


def mark_as_read(message_id):
    resp = requests.patch(
        f"{GRAPH_ROOT}/me/messages/{message_id}",
        headers=graph_headers(),
        json={"isRead": True},
    )
    resp.raise_for_status()


LOG_FILE = os.path.join(DATA_DIR, "marked_read.log")


def log_marked_read(sender, subject):
    line = f"{datetime.now(timezone.utc).isoformat()}\t{sender}\t{subject}\n"
    with open(LOG_FILE, "a") as f:
        f.write(line)


def delete_message(message_id):
    resp = requests.delete(
        f"{GRAPH_ROOT}/me/messages/{message_id}", headers=graph_headers())
    resp.raise_for_status()

# ---------------------------------------------------------------------------
# Flask app
# ---------------------------------------------------------------------------


app_flask = Flask(__name__)


@app_flask.route("/notifications", methods=["GET", "POST"])
def notifications():
    # Graph's validation handshake: it sends validationToken as a query
    # param and expects it echoed back as plain text within 10 seconds.
    validation_token = request.args.get("validationToken")
    if validation_token:
        return Response(validation_token, mimetype="text/plain", status=200)

    data = request.get_json(silent=True) or {}
    for notif in data.get("value", []):
        if notif.get("clientState") != CLIENT_STATE:
            continue  # ignore anything that doesn't match our secret
        message_id = notif["resourceData"]["id"]
        threading.Thread(target=process_new_message,
                         args=(message_id,), daemon=True).start()

    # Acknowledge immediately; Graph expects a fast 202
    return Response(status=202)


def process_new_message(message_id):
    try:
        message = fetch_message(message_id)
        if is_spam(message):
            mark_as_read(message_id)
            sender = message.get("from", {}).get(
                "emailAddress", {}).get("address", "unknown")
            subject = message.get("subject", "(no subject)")
            print(get_time(),
                  f"Spam marked as read: \"{subject}\" from '{sender}'")
            log_marked_read(sender, subject)
        else:
            print(get_time(), f"Left unread: {message.get('subject')}")
    except Exception as e:
        print(get_time(), f"Error processing message {message_id}: {e}")


def get_time():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


if __name__ == "__main__":
    print(get_time(), "app.py starting...")
    os.makedirs(DATA_DIR, exist_ok=True)

    # Graph validates the notification URL while creating the subscription,
    # so the webhook must be listening before that request is made.

    # threading.Thread(
    #     target=app_flask.run,
    #     kwargs={"host": "0.0.0.0", "port": 5000},
    #     daemon=True,
    # ).start()

    threading.Thread(
        target=serve,
        args=(app_flask,),
        kwargs={"host": "0.0.0.0", "port": 5000},
        daemon=True,
    ).start()

    if not NOTIFICATION_URL:
        base_url = wait_for_tunnel_url(TUNNEL_URL_FILE)
        wait_for_tunnel_reachable(base_url)
        NOTIFICATION_URL = f"{base_url}/notifications"

    get_access_token()  # triggers device-code login on first run

    existing_sub = None
    if os.path.exists(SUBSCRIPTION_FILE):
        existing_sub = json.load(open(SUBSCRIPTION_FILE))

    if existing_sub and existing_sub.get("notificationUrl") == NOTIFICATION_URL:
        try:
            print(get_time(), "app.py renewing subscription...")
            renew_subscription(existing_sub["id"])
        except Exception:
            print(get_time(), "app.py creating subscription")
            create_subscription_junkemail()
    else:
        # Either no subscription yet, or the tunnel URL changed since last
        # run (common with quick tunnels) - the old subscription is no
        # longer reachable, so drop it and make a fresh one.
        if existing_sub:
            print(
                get_time(), "Tunnel URL changed since last run, recreating subscription.")
            delete_subscription(existing_sub["id"])
        print(get_time(), "app.py creating subscription (file was not found)")
        create_subscription_junkemail()

    threading.Thread(target=subscription_renewal_loop, daemon=True).start()

    print(get_time(), "Thread created. Flask is running...")
    threading.Event().wait()

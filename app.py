"""
Outlook spam-filter bot using Microsoft Graph change notifications.

Supports one or more mailboxes. Each mailbox ("account") gets its own:
  - device-code login and cached token
  - Graph subscription
  - unique clientState (used to tell incoming notifications apart)

All accounts share one CLIENT_ID (a single Azure app registration can be
consented to by multiple different mailbox owners) and one webhook
endpoint/tunnel URL.

Flow:
  1. For each configured account, authenticate once via device code (MSAL),
     token cached to disk per account.
  2. Create a Graph subscription per account on the Junk folder for
     "created" events.
  3. Run a waitress-served Flask app that:
       - answers the validation handshake Graph sends when creating/renewing
         a subscription
       - receives notifications, works out which account they belong to via
         clientState, fetches the message using that account's token, runs
         spam logic, and marks it read
  4. A background thread renews every account's subscription daily
     (max lifetime for messages is ~4230 minutes / ~2.9 days).

Everything here is free: Graph API calls, MSAL, Flask, waitress. You still
need a public HTTPS URL pointing at this server (see cloudflared notes).
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
import re

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
#    quick-tunnel URL (used automatically ONLY if NOTIFICATION_URL isn't set)
NOTIFICATION_URL = os.environ.get("NOTIFICATION_URL", "").strip() or None
TUNNEL_URL_FILE = os.environ.get("TUNNEL_URL_FILE", "/shared/tunnel_url.txt")

# A secret string used as the base for each account's clientState (Graph
# echoes it back in every notification so you can verify it's really Graph,
# and here it also tells you *which* account the notification is for).
CLIENT_STATE_BASE = os.environ["CLIENT_STATE"]

# Comma-separated list of labels, one per mailbox, e.g. "personal,work".
# Labels are just used for file names and logging - they don't need to
# match the actual email address. Defaults to a single "default" account
# so existing single-mailbox setups keep working unchanged.
ACCOUNT_LABELS = [a.strip() for a in os.environ.get(
    "ACCOUNTS", "default").split(",") if a.strip()]

# All state lives under /data, which is a mounted volume so it survives
# container restarts and can move to a new machine (e.g. the Raspberry Pi).
DATA_DIR = os.environ.get("DATA_DIR", "/data")
GRAPH_ROOT = "https://graph.microsoft.com/v1.0"


def get_time():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# get_access_token() reads-then-writes each account's token cache file, and
# is called concurrently from several places (per-notification handler
# threads, the daily renewal loop, startup). Without serializing access per
# account, two concurrent writes to the same file can interleave and
# corrupt it (surfaces later as a JSON "Extra data" error on load). One
# lock per account label keeps concurrent calls for the SAME account safe
# without blocking different accounts from proceeding in parallel.
_token_locks = {label: threading.Lock() for label in ACCOUNT_LABELS}


def _token_lock(label):
    return _token_locks.setdefault(label, threading.Lock())


def token_cache_file(label):
    return os.path.join(DATA_DIR, f"token_cache_{label}.bin")


def subscription_file(label):
    return os.path.join(DATA_DIR, f"subscription_{label}.json")


def client_state_for(label):
    return f"{CLIENT_STATE_BASE}:{label}"


def label_for_client_state(state):
    for label in ACCOUNT_LABELS:
        if client_state_for(label) == state:
            return label
    return None

# ---------------------------------------------------------------------------
# Auth - one MSAL token cache per account
# ---------------------------------------------------------------------------


def load_token_cache(label):
    cache = msal.SerializableTokenCache()
    path = token_cache_file(label)
    if os.path.exists(path):
        cache.deserialize(open(path, "r").read())
    return cache


def save_token_cache(label, cache):
    if cache.has_state_changed:
        path = token_cache_file(label)
        tmp_path = f"{path}.{os.getpid()}.{threading.get_ident()}.tmp"
        # Write to a temp file first, then atomically rename over the real
        # one. os.replace() is atomic at the filesystem level, so even a
        # concurrent reader or a crash mid-write can never see a partial/
        # corrupted file - it either sees the old complete file or the new
        # complete file, never a mix of both.
        with open(tmp_path, "w") as f:
            f.write(cache.serialize())
        os.replace(tmp_path, path)


def get_access_token(label):
    # Serialize the whole read-modify-write cycle per account, so two
    # concurrent calls for the same mailbox (e.g. a notification handler
    # thread and the daily renewal loop firing at the same moment) can't
    # both load, refresh, and save the cache file at once and corrupt it.
    with _token_lock(label):
        cache = load_token_cache(label)
        app = msal.PublicClientApplication(
            CLIENT_ID, authority=AUTHORITY, token_cache=cache)

        accounts = app.get_accounts()
        result = None
        if accounts:
            result = app.acquire_token_silent(SCOPES, account=accounts[0])

        if not result:
            # First run for this account: interactive device code login
            flow = app.initiate_device_flow(scopes=SCOPES)
            if "user_code" not in flow:
                raise RuntimeError(
                    f"[{label}] Failed to create device flow: {flow}")
            print(
                get_time(), f"[{label}] Sign in with a browser to authenticate this mailbox:")
            print(get_time(), flow["message"])
            result = app.acquire_token_by_device_flow(flow)

        save_token_cache(label, cache)

        if "access_token" not in result:
            raise RuntimeError(f"[{label}] Could not get token: {result}")
        return result["access_token"]


def graph_headers(label):
    return {
        "Authorization": f"Bearer {get_access_token(label)}",
        "Content-Type": "application/json",
    }

# ---------------------------------------------------------------------------
# Tunnel readiness (shared across all accounts - one webhook URL for all)
# ---------------------------------------------------------------------------


def wait_for_tunnel_url(path, timeout=120):
    """Poll the file cloudflared writes its assigned URL to, until it appears.
    Only used when NOTIFICATION_URL is NOT set directly - i.e. quick-tunnel
    mode. If you're on a named tunnel with NOTIFICATION_URL set in .env,
    this function is never called."""
    waited = 0
    last_content = None
    while waited < timeout:
        if os.path.exists(path):
            content = open(path).read().strip()
            if content:
                if content != last_content:
                    print(get_time(),
                          f"Tunnel URL detected from {path}: {content}")
                    last_content = content
                return content
        time.sleep(2)
        waited += 2
    raise RuntimeError(
        f"Timed out after {timeout}s waiting for {path} - is the cloudflared container running?"
    )


def wait_for_tunnel_reachable(base_url, timeout=300):
    """Poll from inside our own container until the URL actually responds,
    so we don't hand Graph a dead/not-yet-connected hostname (which it'll
    reject with a validation error). Runs regardless of whether the URL
    came from .env (named tunnel) or the tunnel file (quick tunnel) -
    both can have a short window before they're actually routable."""
    waited = 0
    print(get_time(), "Checking tunnel reachability:", base_url)
    while waited < timeout:
        try:
            requests.get(base_url, timeout=5)
            print(get_time(), f"Tunnel confirmed reachable: {base_url}")
            return
        except requests.exceptions.RequestException as e:
            print(
                get_time(), f"Not reachable yet ({e.__class__.__name__}), waited {waited}s")
        time.sleep(5)
        waited += 5
    raise RuntimeError(
        f"Tunnel at {base_url} never became reachable after {timeout}s")

# ---------------------------------------------------------------------------
# Subscription management - per account
# ---------------------------------------------------------------------------


def delete_subscription(label, sub_id):
    resp = requests.delete(
        f"{GRAPH_ROOT}/subscriptions/{sub_id}", headers=graph_headers(label))
    if resp.status_code not in (204, 404):
        print(get_time(),
              f"[{label}] Warning: could not delete old subscription {sub_id}: {resp.status_code} {resp.text}")


def create_subscription(label):
    expiration = (datetime.now(timezone.utc) +
                  timedelta(minutes=4200)).isoformat()
    body = {
        "changeType": "created",
        "notificationUrl": NOTIFICATION_URL,
        # "resource": "me/mailFolders('inbox')/messages",
        "resource": "me/mailFolders('junkemail')/messages",
        "expirationDateTime": expiration,
        "clientState": client_state_for(label),
    }
    resp = requests.post(f"{GRAPH_ROOT}/subscriptions",
                         headers=graph_headers(label), json=body)
    if not resp.ok:
        print(get_time(),
              f"[{label}] Subscription creation failed: {resp.status_code} {resp.text}")
    resp.raise_for_status()
    sub = resp.json()
    path = subscription_file(label)
    tmp_path = f"{path}.{os.getpid()}.{threading.get_ident()}.tmp"
    with open(tmp_path, "w") as f:
        json.dump(sub, f)
    os.replace(tmp_path, path)
    print(get_time(
    ), f"[{label}] Subscription created, id={sub['id']}, expires {sub['expirationDateTime']}")
    return sub


def renew_subscription(label, sub_id):
    expiration = (datetime.now(timezone.utc) +
                  timedelta(minutes=4200)).isoformat()
    resp = requests.patch(
        f"{GRAPH_ROOT}/subscriptions/{sub_id}",
        headers=graph_headers(label),
        json={"expirationDateTime": expiration},
    )
    resp.raise_for_status()
    print(get_time(),
          f"[{label}] Subscription {sub_id} renewed until {expiration}")


def ensure_subscription(label):
    """Create, renew, or recreate (if the notification URL changed) this account's subscription."""
    existing_sub = None
    if os.path.exists(subscription_file(label)):
        existing_sub = json.load(open(subscription_file(label)))

    if existing_sub and existing_sub.get("notificationUrl") == NOTIFICATION_URL:
        try:
            renew_subscription(label, existing_sub["id"])
        except Exception:
            create_subscription(label)
    else:
        if existing_sub:
            print(
                get_time(),
                f"[{label}] Notification URL changed since last run "
                f"(was '{existing_sub.get('notificationUrl')}', now '{NOTIFICATION_URL}'), "
                f"recreating subscription."
            )
            delete_subscription(label, existing_sub["id"])
        create_subscription(label)


def subscription_renewal_loop():
    while True:
        time.sleep(60 * 60 * 24)  # check once a day
        for label in ACCOUNT_LABELS:
            try:
                if os.path.exists(subscription_file(label)):
                    sub = json.load(open(subscription_file(label)))
                    renew_subscription(label, sub["id"])
            except Exception as e:
                print(
                    get_time(), f"[{label}] Renewal failed, recreating subscription: {e}")
                try:
                    create_subscription(label)
                except Exception as e2:
                    print(get_time(), f"[{label}] Recreate also failed: {e2}")

# ---------------------------------------------------------------------------
# Spam logic - customize this function
# ---------------------------------------------------------------------------


SPAM_KEYWORDS = ["free money", "act now",
                 "wire transfer", "you have won", "crypto giveaway"]
TRUSTED_DOMAINS = ["microsoft.com"]  # never flag these as spam


def get_header(message, header_name):
    headers = message.get("internetMessageHeaders", []) or []
    for h in headers:
        if h["name"].lower() == header_name.lower():
            return h["value"]
    return None


def is_spam(message, label):
    sender = (message.get("from", {}).get(
        "emailAddress", {}).get("address") or "").lower()
    subject = (message.get("subject") or "").lower()
    body_preview = (message.get("bodyPreview") or "").lower()

    if any(sender.endswith("@" + d) for d in TRUSTED_DOMAINS):
        print(get_time(),
              f"[{label}] Marking following mail as NOT SPAM. Reason: Trusted sender")
        return False

    # detects '@' preceded by exactly 8 uppercase char/digits, followed by 28 char/digits
    real_mail_pattern = re.compile(
        r"^[A-Z0-9]{8}@[A-Z0-9]{28}(\.[a-zA-Z]{1,3})?$")
    if real_mail_pattern.search(sender):
        print(get_time(),
              f"[{label}] Marking following mail as SPAM. Reason: sender not matching email pattern")
        return True

    text = f"{subject} {body_preview}"
    if any(keyword in text for keyword in SPAM_KEYWORDS):
        print(get_time(),
              f"[{label}] Marking following mail as SPAM. Reason: spam keywords detected in subject or body")
        return True

    print(get_time(),
          f"[{label}] Marking following mail as NOT SPAM. Reason: No condition returned True")
    return False

# ---------------------------------------------------------------------------
# Message actions - per account (each needs that account's own token)
# ---------------------------------------------------------------------------


def fetch_message(label, message_id):
    resp = requests.get(
        f"{GRAPH_ROOT}/me/messages/{message_id}"
        "?$select=subject,bodyPreview,from,internetMessageHeaders",
        headers=graph_headers(label),
    )
    resp.raise_for_status()
    return resp.json()


def mark_as_read(label, message_id):
    resp = requests.patch(
        f"{GRAPH_ROOT}/me/messages/{message_id}",
        headers=graph_headers(label),
        json={"isRead": True},
    )
    resp.raise_for_status()


def delete_message(label, message_id):
    resp = requests.delete(
        f"{GRAPH_ROOT}/me/messages/{message_id}", headers=graph_headers(label))
    resp.raise_for_status()


LOG_FILE = os.path.join(DATA_DIR, "marked_read.log")


def log_marked_read(label, sender, subject):
    line = f"{datetime.now(timezone.utc).isoformat()}\t{label}\t{sender}\t{subject}\n"
    with open(LOG_FILE, "a") as f:
        f.write(line)

# ---------------------------------------------------------------------------
# Flask app - single shared webhook endpoint for all accounts
# ---------------------------------------------------------------------------


app_flask = Flask(__name__)


@app_flask.route("/notifications", methods=["GET", "POST"])
def notifications():
    print("\n\n\n\n\n\n")
    print(get_time(), "Notification received:", request.args)
    validation_token = request.args.get("validationToken")
    if validation_token:
        print(get_time(), "validation token")
        return Response(validation_token, mimetype="text/plain", status=200)

    data = request.get_json(silent=True) or {}
    print("DATA RECEIVED FROM NOTIFICATIONS:", data)
    for notif in data.get("value", []):
        label = label_for_client_state(notif.get("clientState"))
        if label is None:
            continue  # doesn't match any known account - ignore
        message_id = notif["resourceData"]["id"]
        threading.Thread(target=process_new_message, args=(
            label, message_id), daemon=True).start()

    return Response(status=202)


def process_new_message(label, message_id):
    try:
        message = fetch_message(label, message_id)
        if is_spam(message, label):
            mark_as_read(label, message_id)
            sender = message.get("from", {}).get(
                "emailAddress", {}).get("address", "unknown")
            subject = message.get("subject", "(no subject)")
            print(get_time(),
                  f"[{label}] Marked as read: \"{subject}\" from {sender}")
            log_marked_read(label, sender, subject)
        else:
            print(get_time(),
                  f"[{label}] Left unread: {message.get('subject')}")
    except Exception as e:
        print(get_time(),
              f"[{label}] Error processing message {message_id}: {e}")


if __name__ == "__main__":
    print(get_time(), f"app.py starting for accounts: {ACCOUNT_LABELS}")
    os.makedirs(DATA_DIR, exist_ok=True)

    # Graph validates the notification URL while creating each subscription,
    # so the webhook must be listening before those requests are made.
    threading.Thread(
        target=serve,
        args=(app_flask,),
        kwargs={"host": "0.0.0.0", "port": 5000},
        daemon=True,
    ).start()

    if NOTIFICATION_URL:
        print(get_time(),
              f"Using NOTIFICATION_URL from environment: {NOTIFICATION_URL}")
    else:
        print(get_time(
        ), f"NOTIFICATION_URL not set - falling back to {TUNNEL_URL_FILE} (quick-tunnel mode)")
        base_url = wait_for_tunnel_url(TUNNEL_URL_FILE)
        NOTIFICATION_URL = f"{base_url}/notifications"

    # Always verify reachability, whether the URL came from .env (named
    # tunnel) or the tunnel file (quick tunnel) - a named tunnel can also
    # take a few seconds to connect after the cloudflared container starts,
    # and skipping this check is exactly what caused the previous
    # "NotFound" validation error.
    wait_for_tunnel_reachable(NOTIFICATION_URL)

    for label in ACCOUNT_LABELS:
        # triggers device-code login on first run, one prompt per account
        get_access_token(label)
        ensure_subscription(label)

    threading.Thread(target=subscription_renewal_loop, daemon=True).start()

    print(get_time(), "All accounts set up. Flask is running...")
    threading.Event().wait()

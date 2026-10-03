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
import traceback

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


def processed_file(label):
    return os.path.join(DATA_DIR, f"processed_messages_{label}.json")


def delta_file(label):
    return os.path.join(DATA_DIR, f"delta_{label}.json")


# Tracks (message_id, folder_id) pairs already handled, per account, so a
# message is only ever acted on once - regardless of how many "created" or
# "updated" notifications Graph later sends for it (our own mark_as_read()
# write triggers one, and so does e.g. a user manually marking a message
# unread again, which should NOT make the bot re-process or re-mark it).
_processed_locks = {label: threading.Lock() for label in ACCOUNT_LABELS}
_processed_cache = {}  # label -> set of "message_id|folder_id" strings


def _processed_lock(label):
    return _processed_locks.setdefault(label, threading.Lock())


def _load_processed(label):
    if label not in _processed_cache:
        path = processed_file(label)
        entries = set()
        if os.path.exists(path):
            try:
                entries = set(json.load(open(path)))
            except (json.JSONDecodeError, OSError):
                entries = set()
        _processed_cache[label] = entries
    return _processed_cache[label]


def _save_processed(label, entries):
    path = processed_file(label)
    tmp_path = f"{path}.{os.getpid()}.{threading.get_ident()}.tmp"
    with open(tmp_path, "w") as f:
        json.dump(sorted(entries), f)
    os.replace(tmp_path, path)


def claim_message(label, message_id, folder_id):
    """Returns True the first time this (message_id, folder_id) pair is
    seen for this account, and records it. Returns False (caller should
    skip) if it was already processed before."""
    key = f"{message_id}|{folder_id}"
    with _processed_lock(label):
        entries = _load_processed(label)
        if key in entries:
            return False
        entries.add(key)
        _save_processed(label, entries)
        return True


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


GRAPH_RETRY_STATUSES = (429, 503, 504)


def graph_request(label, method, url, headers=None, max_attempts=5, **kwargs):
    """requests.request() against Graph, retrying when Graph throttles us
    (429) or is briefly unavailable (503/504). Waits for the Retry-After
    header Graph sends, falling back to exponential backoff. Returns the
    final response - callers still decide what to do with non-2xx codes."""
    for attempt in range(1, max_attempts + 1):
        resp = requests.request(
            method, url, headers={**graph_headers(label), **(headers or {})}, **kwargs)
        if resp.status_code not in GRAPH_RETRY_STATUSES or attempt == max_attempts:
            return resp
        try:
            wait = int(resp.headers.get("Retry-After", ""))
        except ValueError:
            wait = 2 ** attempt
        wait = min(max(wait, 1), 60)
        print(get_time(),
              f"[{label}] Graph returned {resp.status_code}, retrying in {wait}s (attempt {attempt}/{max_attempts})")
        time.sleep(wait)

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


SUBSCRIPTION_RESOURCE = "me/mailFolders('junkemail')/messages"

# "created" alone misses mail that lands in Junk via a post-delivery move
# (e.g. Exchange's Zero-Hour Auto Purge reclassifying a message that was
# already delivered to the Inbox) - a move is reported as "updated", not
# "created", since the message itself isn't new. "updated" also fires on
# our own mark_as_read() write and on a user manually toggling read/unread;
# claim_message() is what keeps those from being reprocessed.
SUBSCRIPTION_CHANGE_TYPE = "created,updated"


def lifecycle_url():
    """Lifecycle notifications (missed / subscriptionRemoved /
    reauthorizationRequired) go to /lifecycle on the same host as
    NOTIFICATION_URL. Computed at call time since NOTIFICATION_URL can be
    filled in at startup from the tunnel file."""
    base = NOTIFICATION_URL.rstrip("/")
    if base.endswith("/notifications"):
        base = base[:-len("/notifications")]
    return f"{base}/lifecycle"


def create_subscription(label):
    expiration = (datetime.now(timezone.utc) +
                  timedelta(minutes=4200)).isoformat()
    body = {
        "changeType": SUBSCRIPTION_CHANGE_TYPE,
        "notificationUrl": NOTIFICATION_URL,
        "lifecycleNotificationUrl": lifecycle_url(),
        "resource": SUBSCRIPTION_RESOURCE,
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

    matches_current_config = (
        existing_sub
        and existing_sub.get("notificationUrl") == NOTIFICATION_URL
        and existing_sub.get("resource") == SUBSCRIPTION_RESOURCE
        and existing_sub.get("changeType") == SUBSCRIPTION_CHANGE_TYPE
        and existing_sub.get("lifecycleNotificationUrl") == lifecycle_url()
    )

    if matches_current_config:
        try:
            renew_subscription(label, existing_sub["id"])
        except Exception:
            create_subscription(label)
    else:
        if existing_sub:
            changes = []
            if existing_sub.get("notificationUrl") != NOTIFICATION_URL:
                changes.append(
                    f"prev url: '{existing_sub.get('notificationUrl')}', now url: '{NOTIFICATION_URL}'")
            if existing_sub.get("resource") != SUBSCRIPTION_RESOURCE:
                changes.append(
                    f"prev resource: '{existing_sub.get('resource')}', now resource: '{SUBSCRIPTION_RESOURCE}'")
            if existing_sub.get("changeType") != SUBSCRIPTION_CHANGE_TYPE:
                changes.append(
                    f"prev changeType: '{existing_sub.get('changeType')}', now changeType: '{SUBSCRIPTION_CHANGE_TYPE}'")
            if existing_sub.get("lifecycleNotificationUrl") != lifecycle_url():
                changes.append(
                    f"prev lifecycle url: '{existing_sub.get('lifecycleNotificationUrl')}', now lifecycle url: '{lifecycle_url()}'")
            print(get_time(
            ), f"[{label}] Subscription config changed ({'; '.join(changes)}), recreating subscription.")
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
# Delta resync - catches changes Graph never sent a notification for
# ---------------------------------------------------------------------------
#
# Graph doesn't guarantee one notification per change: under bursts it
# throttles and drops some, and says so by sending a "missed" lifecycle
# event. A delta query on the Junk folder returns every message that changed
# since the last saved deltaLink, regardless of which notifications actually
# arrived. Each changed message goes through process_new_message(), so
# claim_message() still skips anything push notifications already handled.


DELTA_START_URL = f"{GRAPH_ROOT}/{SUBSCRIPTION_RESOURCE}/delta?$select=parentFolderId"

# A safety-net resync also runs on this interval, since "missed" lifecycle
# events are themselves best-effort. 0 disables it.
RESYNC_INTERVAL_MINUTES = int(os.environ.get("RESYNC_INTERVAL_MINUTES", "15"))

_resync_locks = {label: threading.Lock() for label in ACCOUNT_LABELS}


def _resync_lock(label):
    return _resync_locks.setdefault(label, threading.Lock())


def _load_delta_link(label):
    path = delta_file(label)
    if not os.path.exists(path):
        return None
    try:
        return json.load(open(path)).get("deltaLink")
    except (json.JSONDecodeError, OSError):
        return None


def _save_delta_link(label, delta_link):
    path = delta_file(label)
    tmp_path = f"{path}.{os.getpid()}.{threading.get_ident()}.tmp"
    with open(tmp_path, "w") as f:
        json.dump({"deltaLink": delta_link}, f)
    os.replace(tmp_path, path)


def resync_junk(label, reason):
    """Pull every Junk message that changed since the last resync and process
    them. With no saved delta state (first run), only records a baseline -
    messages already sitting in Junk are not processed, so the bot doesn't
    suddenly act on the whole folder's history."""
    # One resync per account at a time; a second trigger waits and then
    # picks up whatever changed after the first one finished.
    with _resync_lock(label):
        delta_link = _load_delta_link(label)
        baseline = delta_link is None
        url = delta_link or DELTA_START_URL
        restarted = False
        message_ids = []
        new_delta_link = None

        while url:
            resp = graph_request(label, "GET", url, headers={
                                 "Prefer": "odata.maxpagesize=50"})
            if resp.status_code == 410 and not restarted:
                # Saved sync state expired on Graph's side - start over.
                print(get_time(),
                      f"[{label}] Delta state expired, recording a new baseline (changes since the last resync may be skipped)")
                url, baseline, restarted, message_ids = DELTA_START_URL, True, True, []
                continue
            resp.raise_for_status()
            data = resp.json()
            if not baseline:
                message_ids.extend(item["id"] for item in data.get(
                    "value", []) if "@removed" not in item)
            url = data.get("@odata.nextLink")
            new_delta_link = data.get("@odata.deltaLink", new_delta_link)

        if baseline:
            print(get_time(),
                  f"[{label}] Delta baseline recorded ({reason}); existing Junk messages were not processed")
        else:
            message_ids = list(dict.fromkeys(message_ids))
            if message_ids:
                print(get_time(),
                      f"[{label}] Resync ({reason}): {len(message_ids)} changed message(s) to check")
            # Sequential on purpose - spawning a thread per message here is
            # exactly the kind of burst that gets us throttled.
            for message_id in message_ids:
                process_new_message(label, message_id)

        if new_delta_link:
            _save_delta_link(label, new_delta_link)


def periodic_resync_loop():
    while True:
        time.sleep(RESYNC_INTERVAL_MINUTES * 60)
        for label in ACCOUNT_LABELS:
            try:
                resync_junk(label, "periodic")
            except Exception as e:
                print(get_time(), f"[{label}] Periodic resync failed: {e}")

# ---------------------------------------------------------------------------
# Spam logic - customize this function
# ---------------------------------------------------------------------------


SPAM_KEYWORDS = ["free money", "act now",
                 "wire transfer", "you have won", "crypto giveaway"]
TRUSTED_DOMAINS = ["microsoft.com"]  # never flag these as spam

# Each entry is a regex matched against the sender address. Matched
# case-sensitively on purpose - some patterns below (the "noreply@mail....com"
# one) rely on the actual letter case to tell a random spam subdomain apart
# from a real one like mail.anthropic.com, and re.IGNORECASE would defeat
# that. Plain domains are escaped and anchored to "ends with @domain", so
# they behave like the old plain-string list did (case stops mattering for
# those since sender domains are effectively always lowercase anyway).
SPAM_DOMAINS = [
    "@" + re.escape("pridesolutions.nl") + "$",
    "@" + re.escape("in2.getdrip.com") + "$",
    "@" + re.escape("hudzer.com") + "$",
    # Match like: sender@OPTIONAL.origintip.com
    r"@([a-zA-Z0-9.-]*\.)?origintip\.com$",
    # noreply@mail.<11 random mixed-case chars>.com. Requires both an
    # uppercase and a lowercase letter in that segment, so a real,
    # all-lowercase subdomain (mail.anthropic.com) never matches.
    r"^noreply@mail\.(?=[A-Za-z0-9]*[A-Z])(?=[A-Za-z0-9]*[a-z])[A-Za-z0-9]{11}\.com$",
]
SPAM_DOMAIN_PATTERNS = [re.compile(pattern) for pattern in SPAM_DOMAINS]


def get_header(message, header_name):
    headers = message.get("internetMessageHeaders", []) or []
    for h in headers:
        if h["name"].lower() == header_name.lower():
            return h["value"]
    return None


def is_spam(message):
    sender = (message.get("from", {}).get(
        "emailAddress", {}).get("address") or "")
    subject = message.get("subject")

    if any(sender.lower().endswith("@" + d) for d in TRUSTED_DOMAINS):
        return False, "Trusted sender"

    if any(pattern.search(sender) for pattern in SPAM_DOMAIN_PATTERNS):
        return True, "Spam sender"

    # detects '@' preceded by exactly 8 uppercase char/digits, followed by 28 char/digits
    # Used to detect: AAAAAA1A@AAAA8NWA0OS7FZLAAAAAAAAAAAAD.com
    real_mail_pattern = re.compile(
        r"^[A-Z0-9]{8}@[A-Z0-9]{28}(\.[a-zA-Z]{1,3})?$")

    # detects '@' followed by at least one domain character, a literal '.', and 1 to 3 letters until the end of the string
    # Used to detect: no-reply@abcdefosgucovcsw
    missing_domain_extension_pattern = re.compile(
        r"^[^@\s]+@(?![^@\s]+\.[A-Za-z]{1,3}$)[^@\s]+$")

    if real_mail_pattern.search(sender) or missing_domain_extension_pattern.search(sender):
        return True, "Sender not matching email pattern"

    subject = (message.get("subject") or "").lower()
    body_preview = (message.get("bodyPreview") or "").lower()
    text = f"{subject} {body_preview}"

    if any(keyword in text for keyword in SPAM_KEYWORDS):
        return True, "Spam keywords detected in subject or body"

    return False, "No condition returned True"

# ---------------------------------------------------------------------------
# Message actions - per account (each needs that account's own token)
# ---------------------------------------------------------------------------


def fetch_message(label, message_id):
    resp = graph_request(
        label, "GET",
        f"{GRAPH_ROOT}/me/messages/{message_id}"
        "?$select=subject,bodyPreview,from,internetMessageHeaders,parentFolderId",
    )
    resp.raise_for_status()
    return resp.json()


_folder_name_cache = {}


def get_folder_name(label, folder_id):
    key = (label, folder_id)

    if key in _folder_name_cache:
        return _folder_name_cache[key]

    resp = graph_request(
        label, "GET", f"{GRAPH_ROOT}/me/mailFolders/{folder_id}")

    resp.raise_for_status()
    name = resp.json().get("displayName", folder_id)
    _folder_name_cache[key] = name

    return name


def mark_as_read(label, message_id):
    resp = graph_request(
        label, "PATCH", f"{GRAPH_ROOT}/me/messages/{message_id}",
        json={"isRead": True},
    )
    resp.raise_for_status()


def delete_message(label, message_id):
    resp = requests.delete(
        f"{GRAPH_ROOT}/me/messages/{message_id}", headers=graph_headers(label))
    resp.raise_for_status()


LOG_FILE = os.path.join(DATA_DIR, "marked_read.log")


def log_marked(label, sender, subject, is_spam):
    line = f"{datetime.now(timezone.utc).isoformat()}\t{label}\t{'SPAM' if is_spam else 'LEGIT'}\t{sender}\t{subject}\n"
    with open(LOG_FILE, "a") as f:
        f.write(line)

# ---------------------------------------------------------------------------
# Flask app - single shared webhook endpoint for all accounts
# ---------------------------------------------------------------------------


app_flask = Flask(__name__)


@app_flask.route("/notifications", methods=["GET", "POST"])
def notifications():
    validation_token = request.args.get("validationToken")
    if validation_token:
        return Response(validation_token, mimetype="text/plain", status=200)

    data = request.get_json(silent=True) or {}
    for notif in data.get("value", []):
        label = label_for_client_state(notif.get("clientState"))
        if label is None:
            continue  # doesn't match any known account - ignore
        message_id = notif["resourceData"]["id"]
        threading.Thread(target=process_new_message, args=(
            label, message_id), daemon=True).start()

    return Response(status=202)


@app_flask.route("/lifecycle", methods=["GET", "POST"])
def lifecycle():
    validation_token = request.args.get("validationToken")
    if validation_token:
        return Response(validation_token, mimetype="text/plain", status=200)

    data = request.get_json(silent=True) or {}
    for notif in data.get("value", []):
        label = label_for_client_state(notif.get("clientState"))
        if label is None:
            continue  # doesn't match any known account - ignore
        threading.Thread(target=handle_lifecycle_event, args=(
            label, notif.get("lifecycleEvent"), notif.get("subscriptionId")), daemon=True).start()

    return Response(status=202)


def handle_lifecycle_event(label, event, sub_id):
    try:
        print(get_time(), f"[{label}] Lifecycle event received: {event}")
        current_sub_id = None
        if os.path.exists(subscription_file(label)):
            current_sub_id = json.load(
                open(subscription_file(label))).get("id")

        if event == "missed":
            resync_junk(label, "Graph reported missed notifications")
        elif sub_id != current_sub_id:
            # Event for a subscription we've already replaced - nothing to do.
            print(get_time(),
                  f"[{label}] Ignoring '{event}' for old subscription {sub_id}")
        elif event == "reauthorizationRequired":
            try:
                renew_subscription(label, sub_id)
            except Exception:
                create_subscription(label)
        elif event == "subscriptionRemoved":
            create_subscription(label)
            # Anything that arrived while there was no subscription.
            resync_junk(label, "subscription was removed")
    except Exception as e:
        print(
            f"{get_time()} [{label}] Error handling lifecycle event '{event}': {e}\n{traceback.format_exc()}")


last_label_processed = ""


def process_new_message(label, message_id):
    global last_label_processed
    sender = "unknown"
    subject = "(no subject)"
    folder_name = "unknown"

    try:
        message = fetch_message(label, message_id)
        sender = message.get("from", {}).get(
            "emailAddress", {}).get("address", "unknown")
        subject = message.get("subject", "(no subject)")

        folder_id = message.get("parentFolderId")
        if folder_id:
            try:
                folder_name = get_folder_name(label, folder_id)
            except Exception as e:
                print(f"Error when getting the folder's name: {e}")

        # Subscribed to both "created" and "updated" events; this message
        # may have already been handled by an earlier notification (our
        # own mark_as_read() write, or an unrelated update like the user
        # toggling it back to unread). Only act on each (message, folder)
        # once.
        if not claim_message(label, message_id, folder_id):
            return

        spam_condition, reason = is_spam(message)
        time_str = get_time()

        subject_formatted = subject[:14] + \
            '...' if len(subject) > 17 else subject

        if label != last_label_processed:
            last_label_processed = label
            print()

        if spam_condition:
            mark_as_read(label, message_id)
            print(
                f"{time_str} [{label}]-SPAM-: {reason}\n\tSubject: '{subject_formatted}'\n\tFrom: '{sender}'\n\tFolder: '{folder_name}'")
            log_marked(label, sender, subject, True)
        else:
            print(
                f"{time_str} [{label}]-LEGIT-: {reason}\n\tSubject: '{subject_formatted}'\n\tFrom: '{sender}'\n\tFolder: '{folder_name}'")
            log_marked(label, sender, subject, False)

    except Exception as e:
        print(
            f"{get_time()} [{label}] Error processing message '{subject}' from '{sender}': {e}\n{traceback.format_exc()}")


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
        # Catch up on anything that changed while the bot was down (or
        # record the first baseline).
        try:
            resync_junk(label, "startup")
        except Exception as e:
            print(get_time(), f"[{label}] Startup resync failed: {e}")

    threading.Thread(target=subscription_renewal_loop, daemon=True).start()
    if RESYNC_INTERVAL_MINUTES > 0:
        threading.Thread(target=periodic_resync_loop, daemon=True).start()

    print(get_time(), "All accounts set up. Flask is running...")
    threading.Event().wait()

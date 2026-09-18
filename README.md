# Outlook Spam Bot — Graph API Webhook (Docker)

Two containers, run together with `docker compose`:

- **app** — the Python/Flask webhook listener
- **cloudflared** — the free Cloudflare Tunnel that exposes `app` to the internet

Both images are multi-arch (amd64, arm64, armv7), so the exact same
`docker-compose.yml` runs on your dev machine and later on a Raspberry Pi
with no changes.

## 1. Register the app (Azure Portal, free, one-time)

1. https://portal.azure.com → Entra ID → App registrations → New registration.
2. "Accounts in any organizational directory and personal Microsoft
   accounts" works fine for personal use.
3. Authentication → Add platform → "Mobile and desktop applications" →
   check `https://login.microsoftonline.com/common/oauth2/nativeclient`.
4. Authentication → Advanced settings → "Allow public client flows" = Yes.
5. API permissions → Add a permission → Microsoft Graph → Delegated →
   add `Mail.ReadWrite`, `Mail.Read`, `User.Read`, `offline_access`.
6. Copy the "Application (client) ID".

## 2. Configure

```bash
cp .env.example .env
```

Fill in `.env`:

- `CLIENT_ID` — from step 1
- `CLIENT_STATE` — any long random string (generate with
  `python -c "import secrets; print(secrets.token_urlsafe(32))"`)
- `NOTIFICATION_URL` — leave a placeholder for now, you'll get the real
  value in the next step

## 3. First run: get your tunnel URL

```bash
docker compose up
```

Watch the `cloudflared` container's logs — it will print something like:

```
outlook-spam-bot-tunnel  | 2026-09-17 ... INF |  https://random-words-1234.trycloudflare.com |
```

Copy that URL, put it in `.env` as:

```
NOTIFICATION_URL=https://random-words-1234.trycloudflare.com/notifications
```

Then restart so the app picks it up:

```bash
docker compose up -d --force-recreate app
```

## 4. First-run login

Check the `app` container's logs:

```bash
docker compose logs -f app
```

It will print a device-code login prompt like:

```
To sign in, use a web browser to open the page https://microsoft.com/devicelogin
and enter the code XXXXXXXX to authenticate.
```

Do that once in any browser, logged into the mailbox you want to protect.
The resulting token is cached in the `bot-data` Docker volume, so you won't
need to log in again — including after moving to the Raspberry Pi, as long
as you bring that volume with you (step 6).

The app starts listening immediately, then creates its Graph subscription so
Microsoft can complete the notification endpoint validation handshake.

## 5. Stable URL with a named tunnel (recommended before leaving it unattended)

The quick tunnel above is the fastest way to test, but its URL changes
every time the `cloudflared` container restarts — meaning you'd have to
manually update `NOTIFICATION_URL` and recreate the subscription each time.

For something you leave running long-term (like on the Pi), set up a
**named tunnel** instead — still free, just needs a Cloudflare account:

1. Sign up at https://dash.cloudflare.com (free).
2. Install `cloudflared` locally once (or `docker run -it --rm cloudflare/cloudflared:latest tunnel login`) and run `cloudflared tunnel login`.
3. `cloudflared tunnel create outlook-bot` — note the tunnel ID.
4. `cloudflared tunnel route dns outlook-bot bot.yourdomain.com` (requires
   a domain added to your free Cloudflare account — you can also use a
   subdomain of a domain you already manage there).
5. In the Cloudflare dashboard, under the tunnel, generate a **token**.
6. In `docker-compose.yml`, comment out the quick-tunnel `command:` line
   and uncomment the named-tunnel `command`/`environment` lines.
7. Add `CLOUDFLARE_TUNNEL_TOKEN=...` to `.env`.
8. Set `NOTIFICATION_URL=https://bot.yourdomain.com/notifications` in `.env`.
9. `docker compose up -d --force-recreate`

Now the URL never changes across restarts or across machines.

## 6. Moving to the Raspberry Pi

1. Install Docker and the Docker Compose plugin on the Pi (Raspberry Pi OS
   64-bit recommended for best compatibility):
   ```bash
   curl -fsSL https://get.docker.com | sh
   sudo apt install docker-compose-plugin
   ```
2. Copy the whole project folder to the Pi (`scp`, git, USB drive, whatever's easiest).
3. Also copy the `bot-data` volume so you don't have to log in again. Easiest way:
   ```bash
   # on the old machine
   docker run --rm -v outlook-spam-bot_bot-data:/data -v $(pwd):/backup alpine tar czf /backup/bot-data.tar.gz -C /data .
   # copy bot-data.tar.gz to the Pi alongside the project folder, then on the Pi:
   docker volume create outlook-spam-bot_bot-data
   docker run --rm -v outlook-spam-bot_bot-data:/data -v $(pwd):/backup alpine tar xzf /backup/bot-data.tar.gz -C /data
   ```
4. `docker compose up -d` on the Pi. If you're using the named tunnel
   (step 5), nothing else changes — same URL, same subscription. If you're
   still on a quick tunnel, you'll get a new URL and need to update
   `NOTIFICATION_URL` again.

## 7. Customize spam detection

Edit `is_spam()` in `app.py`, then `docker compose up -d --build app` to
apply changes. Ideas, all free:

- Already wired in: reuse Microsoft's own spam confidence score via the
  `X-MS-Exchange-Organization-SCL` header.
- Keyword/regex rules on subject and body preview.
- Check `Authentication-Results` header for SPF/DKIM/DMARC failures.
- Maintain a blocklist of sender domains/addresses.
- Train a local scikit-learn Naive Bayes classifier on spam/ham you
  collect over time — no external API needed.

## Notes / limits

- Graph API calls are free at personal-use volume.
- Mail subscriptions expire after ~2.94 days max; the app renews daily via
  a background thread, so leave the container running.
- The `bot-data` volume holds your token cache and subscription ID — treat
  it like a secret (it's effectively a login session for your mailbox).

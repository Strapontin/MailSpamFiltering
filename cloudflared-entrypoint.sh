#!/bin/sh
set -e

mkdir -p /shared

cloudflared tunnel --no-autoupdate --url http://app:5000 2>&1 | tee /shared/cloudflared.log | \
while IFS= read -r line; do
  echo "$line"
  url=$(echo "$line" | grep -oE 'https://[A-Za-z0-9.-]+\.trycloudflare\.com' || true)
  if [ -n "$url" ]; then
    echo "$url" > /shared/tunnel_url.txt
  fi
done
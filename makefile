rebuild-container:; docker compose up -d --build --no-deps app
compose:; docker compose up -d --build

logs:; docker compose logs -f app
logs-cloudflared:; docker compose logs -f cloudflared

delete-subscriptions:; docker compose exec app sh -c 'rm -fv /data/subscription_*.json'
delete-processed:; docker compose exec app sh -c 'rm -fv /data/processed_messages_*.json'

update-cloudflared:; docker compose pull cloudflared && docker compose up -d --no-deps cloudflared

read-markedread-logs:; docker compose exec app cat /data/marked_read.log
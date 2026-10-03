rebuild-container:; docker compose up -d --build --no-deps app
compose:; docker compose up -d --build

logs:; docker compose logs -f app

delete-subscriptions:; docker compose exec app sh -c 'rm -fv /data/subscription_*.json'
delete-processed:; docker compose exec app sh -c 'rm -fv /data/processed_messages_*.json'
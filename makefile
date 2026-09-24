rebuild-container:; docker compose up -d --build --no-deps app
compose:; docker compose up -d --build

logs:; docker compose logs -f app
# HLS Solution - GitHub + Docker deployment

## One-time: put the code on GitHub (private repo)
1. Copy these files into your project folder (next to main.py): Dockerfile, docker-entrypoint.sh, docker-compose.yml, .gitignore, .dockerignore, .github/, and the new base_v12.py. Your project must also contain extensions/ and web/.
2. `git init && git add . && git commit -m "HLS portal"` then create a PRIVATE repo on github.com and push to branch `main`.
3. Check `git status` shows NO .db, .csv or _hls_secret files. They must stay out.

## One-time: the server (any Linux VPS, or an office server)
1. Install Docker (docs.docker.com/engine/install). Create a folder: `mkdir ~/hls && cd ~/hls`
2. Put docker-compose.yml there, plus a `.env` file:
   GITHUB_REPO=yourname/yourrepo   (lowercase)
   DOMAIN=hls.yourcompany.lk       (point this DNS name at the server; leave out to use plain http on port 80)
3. Log in to GitHub's registry once: `echo YOUR_PAT | docker login ghcr.io -u yourname --password-stdin`
   (PAT = GitHub personal access token with read:packages)
4. Start it: `docker compose up -d`
5. Load your data ONCE:
   `docker compose cp hayleys_redaxis_v12.db hls:/data/hayleys_redaxis_v12.db && docker compose restart hls`
   Then change all default passwords.

## Auto-deploy
In the GitHub repo -> Settings -> Secrets and variables -> Actions, add: SERVER_HOST, SERVER_USER, SERVER_SSH_KEY (private key of a deploy user on the server).
From then on, every push to `main` builds a new image and updates the server. Data in the `hls_data` volume is kept.

## Backups
Everything lives in the Docker volume `hls_data` (database, backups, uploads). Copy it off the server regularly, e.g.
`docker run --rm -v hls_data:/d -v $PWD:/b alpine tar czf /b/hls_$(date +%F).tgz /d`

## Local test
`docker compose up --build` then open http://localhost (or set DOMAIN=:80). 

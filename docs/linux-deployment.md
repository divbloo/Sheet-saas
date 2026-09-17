# Linux Production Deployment

This runbook targets Ubuntu Linux, Nginx, PM2, Node.js 22.12+, and an external
MongoDB deployment such as MongoDB Atlas. Run the commands from the repository
root unless a different directory is shown.

## 1. Before switching servers

1. Take a fresh MongoDB dump and verify that the archive is not empty.
2. Copy the dump, the old server's `backend/.env`, PM2 process list, and the
   deployed Git commit to storage outside the old server. A backup on the same
   failing machine is not sufficient.
3. In MongoDB Atlas, allow the new server's fixed public IP before testing.
4. Lower the DNS TTL before the migration window if the public IP will change.
5. Keep the old server available until login, sheet listing, sheet editing,
   realtime updates, Excel export, and restore-from-version have been tested.

Application data is stored in MongoDB. The Linux host contains application
code, the production build, PM2 state, Nginx configuration, and secrets; it
does not contain the primary sheet database when `MONGO_URI` points to Atlas.

## 2. Provision the host

Install Node.js 22.12 or newer from the official Node.js distribution, then:

```bash
node --version
npm --version
sudo apt update
sudo apt install -y git nginx ufw
sudo npm install -g pm2
```

Create a dedicated, non-login application account and directory:

```bash
sudo adduser --system --group --home /opt/sheet-saas sheet-saas
sudo mkdir -p /opt/sheet-saas/app
sudo chown -R sheet-saas:sheet-saas /opt/sheet-saas
```

Clone the repository using a deploy key or another read-only Git credential:

```bash
sudo -u sheet-saas git clone https://github.com/divbloo/Sheet-saas.git /opt/sheet-saas/app
cd /opt/sheet-saas/app
```

## 3. Configure secrets

Create `/opt/sheet-saas/app/backend/.env` from `backend/.env.example`:

```env
NODE_ENV=production
PORT=5000
MONGO_URI=mongodb+srv://USER:PASSWORD@CLUSTER/sheet-saas
JWT_SECRET=GENERATE_A_RANDOM_SECRET_WITH_AT_LEAST_32_CHARACTERS
FRONTEND_URL=https://sheets.example.com
TRUST_PROXY=1
FORCE_HTTPS=true
PERF_LOG_THRESHOLD_MS=250
```

Protect it and never add it to Git:

```bash
sudo chown sheet-saas:sheet-saas /opt/sheet-saas/app/backend/.env
sudo chmod 600 /opt/sheet-saas/app/backend/.env
```

`FRONTEND_URL` accepts a comma-separated allowlist when more than one exact
origin is required. Do not use `*`. `TRUST_PROXY=1` is correct when exactly one
trusted Nginx proxy is in front of Node.

## 4. Install, verify, and start

```bash
cd /opt/sheet-saas/app
sudo -u sheet-saas npm --prefix backend ci --omit=dev --ignore-scripts
sudo -u sheet-saas npm --prefix frontend ci --include=dev --ignore-scripts
sudo -u sheet-saas npm run verify
sudo -iu sheet-saas sh -lc 'cd /opt/sheet-saas/app && pm2 start ecosystem.config.cjs --env production && pm2 save'
```

Enable PM2 at boot. Run the command printed by `pm2 startup`, making sure it
uses the `sheet-saas` user and `/opt/sheet-saas` home, then save once more:

```bash
sudo env PATH="$PATH:/usr/bin" pm2 startup systemd -u sheet-saas --hp /opt/sheet-saas
sudo -iu sheet-saas pm2 save
```

Verify the application locally:

```bash
curl --fail --silent http://127.0.0.1:5000/healthz
sudo -iu sheet-saas pm2 status
sudo -iu sheet-saas pm2 logs sheet-saas --lines 100 --nostream
```

## 5. Nginx and WebSocket proxy

Create `/etc/nginx/sites-available/sheet-saas`:

```nginx
server {
    listen 80;
    server_name sheets.example.com;

    client_max_body_size 16m;

    location /socket.io/ {
        proxy_pass http://127.0.0.1:5000;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_read_timeout 75s;
    }

    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_http_version 1.1;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

Enable and test it:

```bash
sudo ln -s /etc/nginx/sites-available/sheet-saas /etc/nginx/sites-enabled/sheet-saas
sudo nginx -t
sudo systemctl reload nginx
sudo ufw allow OpenSSH
sudo ufw allow 'Nginx Full'
sudo ufw enable
```

Issue a TLS certificate with the organization's chosen ACME client before
setting `FORCE_HTTPS=true`. Keep port 5000 bound behind the firewall; only
Nginx should be public.

### Cloudflare configuration

When Cloudflare manages the domain:

1. Assign an AWS Elastic IP to the EC2 instance so the origin address does not
   change after a restart.
2. Create an `A` record for the application hostname pointing to that Elastic
   IP and enable **Proxied** (orange cloud).
3. Install either a publicly trusted certificate or a Cloudflare Origin CA
   certificate on Nginx, then select **SSL/TLS > Full (strict)**. Do not use
   Flexible mode.
4. Confirm **Network > WebSockets** is enabled. Cloudflare supports proxied
   WebSocket connections used by Socket.IO.
5. Enable automatic HTTPS redirects only after the origin certificate and the
   Nginx HTTPS server block are working.
6. Restrict the EC2 security group ports 80/443 to Cloudflare's published IP
   ranges, or configure Authenticated Origin Pulls, so clients cannot bypass
   Cloudflare and spoof forwarded IP headers.
7. Configure Nginx's real-IP module with Cloudflare's current published address
   ranges and `CF-Connecting-IP` before relying on IP-based rate limiting.

Official references:

- <https://developers.cloudflare.com/dns/proxy-status/>
- <https://developers.cloudflare.com/ssl/origin-configuration/ssl-modes/full-strict/>
- <https://developers.cloudflare.com/ssl/origin-configuration/origin-ca/>
- <https://developers.cloudflare.com/network/websockets/>
- <https://developers.cloudflare.com/ssl/origin-configuration/authenticated-origin-pull/>

## 6. Repeatable deployment

Create a database backup first, then deploy only a known commit from `main`:

```bash
cd /opt/sheet-saas/app
sudo -u sheet-saas git pull --ff-only origin main
sudo -u sheet-saas npm --prefix backend ci --omit=dev --ignore-scripts
sudo -u sheet-saas npm --prefix frontend ci --include=dev --ignore-scripts
sudo -u sheet-saas npm run verify
sudo -iu sheet-saas sh -lc 'cd /opt/sheet-saas/app && pm2 reload ecosystem.config.cjs --env production --update-env && pm2 save'
curl --fail --silent https://sheets.example.com/healthz
git rev-parse --short HEAD
```

PM2 waits for the process to connect to MongoDB and announce readiness before
completing a reload. SIGTERM triggers a graceful Socket.IO, HTTP, and MongoDB
shutdown.

## 7. MongoDB backup and restore

Install MongoDB Database Tools. Store backups outside the application checkout
and copy them to a second machine or object-storage bucket:

```bash
sudo install -d -m 700 -o sheet-saas -g sheet-saas /var/backups/sheet-saas
sudo -iu sheet-saas sh -lc 'set -a; . /opt/sheet-saas/app/backend/.env; set +a; mongodump --uri="$MONGO_URI" --archive="/var/backups/sheet-saas/mongo-$(date +%Y%m%d-%H%M%S).archive.gz" --gzip'
sudo -u sheet-saas find /var/backups/sheet-saas -type f -size +0 -name '*.archive.gz' -ls
```

Test restores into a separate database or Atlas test project first. The
following command can overwrite data when `--drop` is used; verify the target
URI and archive path before running it:

```bash
mongorestore --uri="$RESTORE_MONGO_URI" --archive=/secure/path/mongo-TIMESTAMP.archive.gz --gzip --drop
```

## 8. Rollback

Record the current commit before every deployment. If code must be rolled back:

```bash
cd /opt/sheet-saas/app
git log --oneline -5
sudo -u sheet-saas git switch --detach PREVIOUS_COMMIT_OR_TAG
sudo -u sheet-saas npm --prefix backend ci --omit=dev --ignore-scripts
sudo -u sheet-saas npm --prefix frontend ci --include=dev --ignore-scripts
sudo -u sheet-saas npm run build
sudo -iu sheet-saas sh -lc 'cd /opt/sheet-saas/app && pm2 reload ecosystem.config.cjs --env production --update-env'
curl --fail --silent https://sheets.example.com/healthz
```

Do not restore the database for a code-only rollback unless a verified data
migration explicitly requires it.

## 9. Post-deployment smoke test

- `GET /healthz` returns 200.
- Login succeeds and the sheet list loads.
- A large sheet opens, scrolls, saves, and reloads correctly.
- Two browsers receive realtime cell and presence updates.
- Excel `.xlsx`, CSV, and PDF exports download successfully.
- PM2 shows one online process and logs contain no restart loop.
- Nginx access/error logs show no repeated 4xx/5xx spike.

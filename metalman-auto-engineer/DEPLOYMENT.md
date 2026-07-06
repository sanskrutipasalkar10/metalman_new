# Deploying to an Oracle Cloud VM

This app ships as two Docker images (`backend`, `frontend`) wired together
with `docker-compose.yml` at the repo root. Frontend and backend run on the
same VM; Nginx (inside the `frontend` container) serves the built React app
and reverse-proxies `/api` and `/outputs` to the backend, so only one public
port (80) is needed.

## 1. Provision the VM

- Oracle Cloud → Compute → Create Instance. Ubuntu 22.04+ recommended.
  A "VM.Standard.E4.Flex" with 2-4 OCPUs / 8-16 GB RAM is enough for this
  workload (CAD tessellation + Excel generation are the heaviest steps).
- **Open ports** in the instance's attached Security List / Network Security
  Group: TCP 80 (HTTP), and 443 if you add TLS later. SSH (22) is usually
  already open.
- Also run, on the VM itself, `sudo ufw allow 80/tcp` (or the equivalent) if
  `ufw`/`iptables` is enabled — OCI's cloud firewall and the instance's OS
  firewall are both in the path.

## 2. Install Docker

```bash
curl -fsSL https://get.docker.com | sudo sh
sudo usermod -aG docker $USER
newgrp docker
docker compose version   # confirm the Compose plugin is present
```

## 3. Get the code and configure secrets

```bash
git clone <your-repo-url>
cd metalman_app/metalman-auto-engineer/backend
cp .env.example .env
nano .env   # set JWT_SECRET_KEY, and either AI_PROVIDER=ollama or =gemini + GEMINI_API_KEY
cd ../..
```

Every env var the backend reads is documented in
[`backend/.env.example`](metalman-auto-engineer/backend/.env.example).
At minimum for a cloud deploy:

- `JWT_SECRET_KEY` — generate with `openssl rand -hex 32`.
- `AI_PROVIDER=gemini` and `GEMINI_API_KEY=...` — Ollama Cloud is typically
  not reachable from a bare VM; Gemini is the deployment-friendly path.
  (Leave `AI_PROVIDER=ollama` if you're running Ollama locally/on the VM
  itself and it's reachable at `OLLAMA_BASE_URL`.)

## 4. Build and run

```bash
docker compose up -d --build
docker compose ps          # both services should be "running"/"healthy"
```

Visit `http://<vm-public-ip>/` — you should see the dashboard. The frontend
talks to the backend through the Nginx proxy on the same origin, so no CORS
configuration or separate backend URL is needed for this topology.

## 5. Persistent data

Uploads, generated outputs, and the SQLite database all live under `/data`
inside the `backend` container, which is backed by the named Docker volume
`backend_data` (see `docker-compose.yml`). This survives
`docker compose down` / VM reboots; it's only lost if you explicitly run
`docker compose down -v`.

To back it up:

```bash
docker run --rm -v metalman-auto-engineer_backend_data:/data -v $(pwd):/backup \
  alpine tar czf /backup/metalman-data-backup.tar.gz -C / data
```

## 6. Updating after a code change

```bash
git pull
docker compose up -d --build
```

## Known limitations on Linux

- **Excel → PDF preview** (`/api/preview/{filename}`) uses native Excel COM
  automation on Windows for the highest-fidelity output. That's unavailable
  on Linux, so the VM deployment automatically falls back to the built-in
  lower-fidelity HTML data preview. Both paths already exist in
  `api/routes.py` — this is a graceful, working fallback, not an error.
- Run `docker compose up --build` locally first to catch any issues before
  deploying to the VM — it's the same image, just faster to iterate on.

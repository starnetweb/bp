# Deployment Guide — Hostinger VPS

## Files Included

- **Dockerfile** — Containerizes the Flask application
- **docker-compose.yml** — Orchestrates the container (required by hPanel)
- **.dockerignore** — Excludes unnecessary files from the image
- **requirements.txt** — Python dependencies

## Prerequisites

1. **Hostinger VPS** with Docker installed
2. **Git** installed on the VPS
3. **Anthropic API key** (will be added to config.py)

## Deployment Steps

### Step 1: Clone Repository on VPS

```bash
cd /home/yourusername
git clone https://github.com/starnetweb/bp.git
cd bp
```

### Step 2: Configure API Key

Edit `config.py` and add your Anthropic API key:

```bash
nano config.py
```

Replace:
```python
ANTHROPIC_API_KEY = "sk-ant-api03-..."
```

Save and exit (Ctrl+X, then Y, then Enter).

### Step 3: Deploy with Docker Compose

```bash
# Build the image
docker-compose build

# Start the container
docker-compose up -d

# View logs
docker-compose logs -f app
```

### Step 4: Configure hPanel (If Using hPanel)

In hPanel Docker section:
1. Navigate to **Containers** or **Docker**
2. Click **Deploy from Compose File**
3. Paste the contents of `docker-compose.yml`
4. Set the Docker Compose file location or upload it
5. Click **Deploy**

The app will be available on port **5000**.

## Accessing the App

If your VPS IP is `123.45.67.89`:

```
http://123.45.67.89:5000
```

## Production Setup (Reverse Proxy)

For production, use **Nginx** as a reverse proxy:

```nginx
server {
    listen 80;
    server_name yourdomain.com;

    location / {
        proxy_pass http://localhost:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_connect_timeout 60s;
        proxy_send_timeout 60s;
        proxy_read_timeout 60s;
    }
}
```

## Stopping the Container

```bash
docker-compose down
```

## Viewing Logs

```bash
# Real-time logs
docker-compose logs -f app

# Last 100 lines
docker-compose logs --tail=100 app
```

## Updating the App

```bash
# Pull latest changes
git pull origin main

# Rebuild and restart
docker-compose build
docker-compose down
docker-compose up -d
```

## Troubleshooting

### Port Already in Use
If port 5000 is already in use, edit `docker-compose.yml`:
```yaml
ports:
  - "8080:5000"  # Map to 8080 instead
```

### Container Won't Start
```bash
docker-compose logs app  # Check error logs
```

### Memory Issues
Adjust in `docker-compose.yml`:
```yaml
deploy:
  resources:
    limits:
      cpus: '1'      # Reduce from 2
      memory: 1G     # Reduce from 2G
```

### Generated Files Not Persisting
Make sure the `downloads/` directory is writable:
```bash
chmod 777 downloads/
```

## Security Notes

1. **Never commit `config.py` with real API keys** — use environment variables in production:
   ```yaml
   environment:
     - ANTHROPIC_API_KEY=${ANTHROPIC_API_KEY}
   ```

2. **Use HTTPS** — set up Let's Encrypt with Nginx

3. **Limit API access** — add rate limiting in Nginx

## Support

For issues, check:
- Container logs: `docker-compose logs app`
- VPS system logs: `/var/log/syslog`
- Docker system status: `docker system df`

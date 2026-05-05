# Hostinger hPanel Quick Start

## What to Do in hPanel

### Option 1: Upload Docker Compose File

1. In hPanel, go to **Services → Docker** (or similar)
2. Click **Add Container** or **Deploy from Compose**
3. **Copy-paste this entire block** into the Docker Compose field:

```yaml
version: '3.8'

services:
  app:
    build:
      context: .
      dockerfile: Dockerfile
    container_name: research-writeup-app
    ports:
      - "5000:5000"
    environment:
      - FLASK_ENV=production
      - PYTHONUNBUFFERED=1
    volumes:
      - ./downloads:/app/downloads
      - ./config.py:/app/config.py:ro
    restart: always
    healthcheck:
      test: ["CMD", "curl", "-f", "http://localhost:5000/"]
      interval: 30s
      timeout: 10s
      retries: 3
      start_period: 5s
    deploy:
      resources:
        limits:
          cpus: '2'
          memory: 2G
        reservations:
          cpus: '1'
          memory: 1G
```

4. Click **Deploy**

### Option 2: Connect Git Repository

1. In hPanel Docker settings, find **Git Integration** or **Clone from Repository**
2. Enter: `https://github.com/starnetweb/bp.git`
3. hPanel will auto-detect the `docker-compose.yml`
4. Click **Deploy**

## Before Deploying

✅ Make sure `config.py` has your Anthropic API key:
```python
ANTHROPIC_API_KEY = "sk-ant-api03-YOUR_ACTUAL_KEY_HERE"
```

## After Deployment

Your app will be available at:
```
http://your-vps-ip:5000
```

Or if you set up Nginx reverse proxy:
```
http://yourdomain.com
```

## Check if It's Running

```bash
# SSH into your VPS and run:
docker-compose logs -f app
```

You should see:
```
[2026-05-05 12:00:00 +0000] [1] [INFO] Starting gunicorn 25.3.0
[2026-05-05 12:00:00 +0000] [1] [INFO] Listening at: http://0.0.0.0:5000
```

## Stop/Restart

```bash
docker-compose down      # Stop
docker-compose up -d     # Start
docker-compose restart   # Restart
```

That's it! The app is now running on your VPS.

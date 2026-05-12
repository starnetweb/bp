# Environment Variables Setup Guide

## Overview

The application uses environment variables to manage sensitive credentials (API keys, email passwords, etc.). This keeps secrets out of the codebase and makes it easy to deploy to different environments (local, Hostinger, etc.).

## Local Development Setup

### Step 1: Install python-dotenv
```bash
pip install python-dotenv
```

Or if already installed via requirements.txt:
```bash
pip install -r requirements.txt
```

### Step 2: Create .env file
```bash
cp .env.example .env
```

### Step 3: Edit .env with your credentials
```bash
nano .env    # or use your preferred editor
```

**Required fields:**
- `ANTHROPIC_API_KEY` - Your Claude API key from https://console.anthropic.com/settings/keys
- `SMTP_USER` - Your Gmail address (or SMTP email)
- `SMTP_PASSWORD` - Gmail App Password (NOT your regular password)

**Optional fields:**
- `MODEL` - Default: `claude-sonnet-4-6` (change for cheaper/faster options)
- `SMTP_HOST` - Default: `smtp.gmail.com` (change for non-Gmail)
- `SMTP_PORT` - Default: `587`
- `RECIPIENT_EMAIL_1`, `RECIPIENT_EMAIL_2` - Recipient addresses

### Step 4: Run the application
```bash
python web_app.py
```

The app will automatically load variables from `.env`.

---

## Hostinger Deployment Setup

### Important: Do NOT commit .env to git
The `.env` file is in `.gitignore` and should never be committed. Only `.env.example` is tracked in git.

### Option 1: SSH Terminal (Recommended)

1. **Connect via SSH** to your Hostinger account
2. **Create .env file:**
   ```bash
   cd public_html/Claude_code/p
   cp .env.example .env
   nano .env
   ```
3. **Add your credentials** to .env
4. **Save and exit** (Ctrl+X, then Y, then Enter if using nano)

### Option 2: Hostinger Control Panel (If available)

1. **Go to: Hostinger Dashboard → Environment Variables** (or similar section)
2. **Add each variable:**
   - Name: `ANTHROPIC_API_KEY`
   - Value: `sk-ant-...`
   
   Repeat for:
   - `SMTP_USER`
   - `SMTP_PASSWORD`
   - `MODEL`
   - `RECIPIENT_EMAIL_1`
   - `RECIPIENT_EMAIL_2`

3. **Save** and restart the application

### Option 3: Deploy Script

If deploying via a script or CI/CD pipeline, you can set variables during deployment:

```bash
#!/bin/bash
cd /path/to/Claude_code/p

# Create .env from template
cp .env.example .env

# Set variables (replace with actual values)
echo "ANTHROPIC_API_KEY=${ANTHROPIC_API_KEY}" >> .env
echo "SMTP_USER=${SMTP_USER}" >> .env
echo "SMTP_PASSWORD=${SMTP_PASSWORD}" >> .env

# Start the app
python web_app.py
```

Then set the variables in Hostinger's deployment settings.

---

## Checking Configuration

### Verify .env is loaded
The app will print a warning if `ANTHROPIC_API_KEY` is not set:
```
⚠ WARNING: ANTHROPIC_API_KEY not set in .env file
```

### View current configuration
```python
# In Python shell:
from config import *
print(f"API Key set: {bool(ANTHROPIC_API_KEY)}")
print(f"Model: {MODEL}")
print(f"SMTP Host: {SMTP_HOST}")
```

---

## Priority Order

The application loads configuration in this order (first match wins):

1. **Environment variables** (from .env file or system environment)
2. **Fallback defaults** in config.py

Example:
```python
# This means:
# 1. Check for ANTHROPIC_API_KEY env var
# 2. If not found, use empty string
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
```

---

## Machine-Specific Overrides (.env.local)

For local development with different settings:

```bash
cp .env .env.local
# Edit .env.local with machine-specific values
```

Both `.env` and `.env.local` are gitignored.

---

## Gmail Setup (Recommended)

If using Gmail as your email provider:

1. **Enable 2-Step Verification** in your Google Account
2. **Generate App Password:**
   - Go to https://myaccount.google.com/apppasswords
   - Select: Mail → Windows Computer (or your device)
   - Copy the 16-character password
3. **Add to .env:**
   ```
   SMTP_USER=your_email@gmail.com
   SMTP_PASSWORD=xxxx xxxx xxxx xxxx
   ```

**Do NOT use your regular Google password** — use only the App Password.

---

## Troubleshooting

### "ANTHROPIC_API_KEY not set"
- Check that .env file exists in the project root
- Verify the API key is not blank: `ANTHROPIC_API_KEY=sk-ant-...`
- Restart the application after creating/editing .env

### "Email failed: Authentication error"
- Verify SMTP_USER and SMTP_PASSWORD are correct
- For Gmail, ensure you're using an App Password (not your login password)
- Check that 2-Step Verification is enabled on your Google Account

### ".env not being loaded on Hostinger"
- Verify the file exists: `ls -la .env`
- Check file permissions: `chmod 644 .env`
- Ensure python-dotenv is installed: `pip list | grep python-dotenv`
- Restart the application/web server

---

## Security Best Practices

✅ **DO:**
- Keep .env out of git (it's in .gitignore)
- Use environment variables for all sensitive data
- Rotate API keys periodically
- Use Gmail App Passwords (not main password)
- Restrict .env file permissions: `chmod 600 .env`

❌ **DON'T:**
- Commit .env to git
- Share .env files via email or chat
- Hardcode credentials in Python files
- Use the same API key across multiple projects
- Leave .env readable by other users

---

## Reference

| Variable | Purpose | Example | Required |
|----------|---------|---------|----------|
| `ANTHROPIC_API_KEY` | Claude API key | `sk-ant-...` | ✅ Yes |
| `MODEL` | Claude model to use | `claude-sonnet-4-6` | ❌ No |
| `SMTP_HOST` | Email server address | `smtp.gmail.com` | ❌ No |
| `SMTP_PORT` | Email server port | `587` | ❌ No |
| `SMTP_USER` | Email sender address | `your@gmail.com` | ❌ No |
| `SMTP_PASSWORD` | Email server password | `xxxx xxxx xxxx xxxx` | ❌ No |
| `RECIPIENT_EMAIL_1` | Primary recipient email | `recipient@example.com` | ❌ No |
| `RECIPIENT_EMAIL_2` | Secondary recipient email | `recipient@example.com` | ❌ No |

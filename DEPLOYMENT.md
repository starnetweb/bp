# Deployment Setup Guide

## Configuration

The application requires a `config.py` file that is **not** committed to the repository for security reasons (API keys are sensitive).

### Setup Steps

1. **Copy the template:**
   ```bash
   cp config.py.example config.py
   ```

2. **Edit `config.py` with your values:**
   ```python
   # Get Anthropic API key from: https://console.anthropic.com/settings/keys
   ANTHROPIC_API_KEY = "your-actual-anthropic-api-key"
   
   # Choose model (default is fine)
   MODEL = "claude-sonnet-4-6"
   
   # Email settings (Gmail or your provider)
   SMTP_HOST = "smtp.gmail.com"
   SMTP_PORT = 587
   SMTP_USER = "your-email@gmail.com"
   SMTP_PASSWORD = "your-app-password"
   
   # Recipients for generated documents
   RECIPIENT_EMAILS = [
       "recipient1@example.com",
       "recipient2@example.com",
   ]
   ```

3. **For Gmail:**
   - Enable 2-factor authentication on your Gmail account
   - Generate an **App Password** (NOT your regular password)
   - Use the App Password in `SMTP_PASSWORD`
   - See: https://support.google.com/accounts/answer/185833

4. **Verify locally before deployment:**
   ```bash
   python web_app.py
   # Visit http://localhost:5000 in your browser
   ```

5. **In production/container deployment:**
   - Ensure `config.py` exists in the application directory before starting
   - The file should have proper permissions (readable by the application user)
   - Never commit `config.py` to version control

## Security Notes

- `config.py` is in `.gitignore` — it is NOT tracked by git
- Never share or expose the contents of `config.py`
- Rotate API keys periodically
- Use environment variables in sensitive environments if possible

## Troubleshooting

### `ModuleNotFoundError: No module named 'config'`

This error means `config.py` is missing. Run:
```bash
cp config.py.example config.py
# Then edit config.py with your actual values
```

### App won't start in Docker/container

Make sure `config.py` is copied into the container before the app starts.

## Quick Fix for Current Deployment

Your app is running in a container at `/app/`. You need to create `config.py` in that location:

```bash
cd /app
cp config.py.example config.py
# Then edit config.py with your actual credentials
nano config.py
```

After saving, restart the application.

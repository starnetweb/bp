# -*- coding: utf-8 -*-
# =============================================================
#  CONFIGURATION — Environment Variables & Defaults
# =============================================================

import os
from dotenv import load_dotenv

# Load environment variables from .env file in the same directory as this config.py
dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
load_dotenv(dotenv_path=dotenv_path)

# ── Anthropic API ─────────────────────────────────────────────
# Get your key at: https://console.anthropic.com/settings/keys
# Load from .env or use empty string
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")

if not ANTHROPIC_API_KEY:
    print("⚠ WARNING: ANTHROPIC_API_KEY not set in .env file")

# ── Model selection ───────────────────────────────────────────
# Options (best → cheapest):
#   "claude-opus-4-6"    ~$0.40–0.80 per document  (best quality)
#   "claude-sonnet-4-6"  ~$0.10–0.20 per document  (great quality)
#   "claude-haiku-4-5"   ~$0.03–0.06 per document  (good for drafts)
MODEL = os.getenv("MODEL", "claude-sonnet-4-6")

# ── Email delivery ────────────────────────────────────────────
# The finished .docx will be emailed to addresses configured below.
# Uses Gmail by default — for other providers adjust SMTP_HOST/PORT.

SMTP_HOST     = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT     = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER     = os.getenv("SMTP_USER", "youremail@gmail.com")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "your-gmail-app-password")

# Build recipient list from environment variables
RECIPIENT_EMAILS = []
email1 = os.getenv("RECIPIENT_EMAIL_1", "")
email2 = os.getenv("RECIPIENT_EMAIL_2", "")

if email1:
    RECIPIENT_EMAILS.append(email1)
if email2:
    RECIPIENT_EMAILS.append(email2)

# Fallback to defaults if no env vars set
if not RECIPIENT_EMAILS:
    RECIPIENT_EMAILS = [
        "recipient1@example.com",
        "recipient2@example.com",
    ]

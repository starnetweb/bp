# -*- coding: utf-8 -*-
# =============================================================
#  CONFIGURATION — Settings
#  This file is generated from config.py.example
#  Values can be overridden by environment variables
# =============================================================

import os

# ── Anthropic API ─────────────────────────────────────────────
# Get your key at: https://console.anthropic.com/settings/keys
# Can be set via ANTHROPIC_API_KEY environment variable
ANTHROPIC_API_KEY = ""

# ── Model selection ───────────────────────────────────────────
# Options (best → cheapest):
#   "claude-opus-5-5"    ~$0.30–0.60 per document  (best quality)
#   "claude-sonnet-5-5"  ~$0.07–0.15 per document  (great quality, recommended)
#   "claude-haiku-5-5"   ~$0.02–0.04 per document  (good for drafts)
MODEL = "claude-sonnet-5-5"

# ── Email delivery ────────────────────────────────────────────
# The finished .docx will be emailed to these addresses.
# Uses Gmail by default — for other providers adjust SMTP_HOST/PORT.

SMTP_HOST     = "smtp.gmail.com"
SMTP_PORT     = 587
SMTP_USER     = "your_email@gmail.com"
SMTP_PASSWORD = "your-gmail-app-password"

RECIPIENT_EMAILS = [
    "recipient1@example.com",
    "recipient2@example.com",
]

# -*- coding: utf-8 -*-
# =============================================================
#  CONFIGURATION — Settings
# =============================================================

# ── Anthropic API ─────────────────────────────────────────────
# Get your key at: https://console.anthropic.com/settings/keys
ANTHROPIC_API_KEY = "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"

# ── Model selection ───────────────────────────────────────────
# Options (best → cheapest):
#   "claude-opus-4-6"    ~$0.40–0.80 per document  (best quality)
#   "claude-sonnet-4-6"  ~$0.10–0.20 per document  (great quality, recommended)
#   "claude-haiku-4-5"   ~$0.03–0.06 per document  (good for drafts)
MODEL = "claude-sonnet-4-6"

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

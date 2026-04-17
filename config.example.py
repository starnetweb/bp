# =============================================================
#  CONFIGURATION — copy this file to config.py and fill in values
# =============================================================

# ── Anthropic API ─────────────────────────────────────────────
ANTHROPIC_API_KEY = "sk-ant-api03-..."

# ── Model selection ───────────────────────────────────────────
# Options (best → cheapest):
#   "claude-opus-4-6"    (best quality)
#   "claude-sonnet-4-6"  (great quality, recommended)
#   "claude-haiku-4-5"   (good for drafts)
MODEL = "claude-sonnet-4-6"

# ── Email delivery ────────────────────────────────────────────
SMTP_HOST     = "smtp.gmail.com"
SMTP_PORT     = 587
SMTP_USER     = "youremail@gmail.com"
SMTP_PASSWORD = "your-gmail-app-password"

RECIPIENT_EMAILS = [
    "recipient1@example.com",
    "recipient2@example.com",
]

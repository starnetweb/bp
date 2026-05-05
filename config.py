# =============================================================
#  CONFIGURATION — edit this file to change settings
# =============================================================

# ── Anthropic API ─────────────────────────────────────────────
# Get your key at: https://console.anthropic.com/settings/keys
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")

# ── Model selection ───────────────────────────────────────────
# Options (best → cheapest):
#   "claude-opus-4-6"    ~$0.40–0.80 per document  (best quality)
#   "claude-sonnet-4-6"  ~$0.10–0.20 per document  (great quality)
#   "claude-haiku-4-5"   ~$0.03–0.06 per document  (good for drafts)
MODEL = "claude-sonnet-4-6"

# ── Email delivery ────────────────────────────────────────────
# The finished .docx will be emailed to BOTH addresses below.
# Uses Gmail by default — for other providers adjust SMTP_HOST/PORT.

SMTP_HOST     = "smtp.gmail.com"
SMTP_PORT     = 587
SMTP_USER     = "youremail@gmail.com"       # sender Gmail address
SMTP_PASSWORD = "your-gmail-app-password"   # Gmail App Password (not your login password)
                                             # Generate at: myaccount.google.com/apppasswords

RECIPIENT_EMAILS = [
    "recipient1@example.com",
    "recipient2@example.com",
]

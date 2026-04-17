<?php
// ─────────────────────────────────────────────────────────
//  RESEARCH PORTAL — CONFIGURATION
//  Edit these values before deployment
// ─────────────────────────────────────────────────────────

// Database
define('DB_HOST', 'localhost');
define('DB_NAME', 'research_portal');
define('DB_USER', 'root');
define('DB_PASS', '');

// Flask / Python research-agent API
define('FLASK_API_URL', 'http://localhost:5000/generate');

// Admin login password
define('ADMIN_PASSWORD', 'admin@2024');

// Session name
define('SESSION_NAME', 'rp_admin');

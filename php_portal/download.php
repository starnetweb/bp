<?php
/**
 * download.php — Admin proxy to download a completed research document from the Flask server.
 *
 * Usage: download.php?id={topic_row_id}
 * Requires admin session.
 */
require_once __DIR__ . '/db.php';
require_once __DIR__ . '/config.php';

session_name(SESSION_NAME);
session_start();

// ── Auth ──────────────────────────────────────────────────
if (empty($_SESSION['admin'])) {
    http_response_code(401);
    die('Unauthorised — please log in to the admin panel first.');
}

// ── Validate input ────────────────────────────────────────
$id = (int)(isset($_GET['id']) ? $_GET['id'] : 0);
if (!$id) {
    http_response_code(400);
    die('Invalid topic ID.');
}

// ── Fetch row from DB ─────────────────────────────────────
$db   = getDB();
$stmt = $db->prepare('SELECT id, topic, job_id, status FROM topics WHERE id = ?');
$stmt->execute(array($id));
$row = $stmt->fetch();

if (!$row) {
    http_response_code(404);
    die('Topic not found.');
}

if (empty($row['job_id'])) {
    http_response_code(404);
    die('This topic has not been approved yet — no document exists.');
}

// ── Build Flask download URL ──────────────────────────────
// Strip "/generate" from the end of FLASK_API_URL to get the base URL
$flask_base   = rtrim(preg_replace('#/generate$#', '', FLASK_API_URL), '/');
$download_url = $flask_base . '/download/' . rawurlencode($row['job_id']);

// ── Fetch document from Flask ─────────────────────────────
$ch = curl_init($download_url);
curl_setopt_array($ch, array(
    CURLOPT_RETURNTRANSFER => true,
    CURLOPT_FOLLOWLOCATION => true,
    CURLOPT_TIMEOUT        => 60,
    CURLOPT_CONNECTTIMEOUT => 10,
    CURLOPT_HEADER         => true,
));
$response = curl_exec($ch);
$httpCode  = curl_getinfo($ch, CURLINFO_HTTP_CODE);
$headerSz  = curl_getinfo($ch, CURLINFO_HEADER_SIZE);
$curlErr   = curl_error($ch);
curl_close($ch);

if ($curlErr) {
    http_response_code(502);
    die('Could not reach the document server: ' . $curlErr);
}

if ($httpCode === 404) {
    http_response_code(404);
    die('Document not found on the server. It may have been cleared after a server restart. Please re-approve the topic to regenerate it.');
}

if ($httpCode !== 200) {
    http_response_code(502);
    die("Document server returned HTTP $httpCode.");
}

// ── Stream response to browser ────────────────────────────
$body = substr($response, $headerSz);

// Build a safe filename from the topic
$safe     = preg_replace('/[^\w\s\-]/', '', $row['topic']);
$safe     = trim(str_replace(' ', '_', $safe));
$safe     = substr($safe, 0, 60);
$filename = 'Research_' . $safe . '.docx';

header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
header('Content-Disposition: attachment; filename="' . $filename . '"');
header('Content-Length: ' . strlen($body));
header('Cache-Control: no-cache, no-store');
echo $body;
exit;

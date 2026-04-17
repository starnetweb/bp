<?php
require_once __DIR__ . '/db.php';
require_once __DIR__ . '/config.php';

header('Content-Type: application/json');
session_name(SESSION_NAME);
session_start();

// Auth guard
if (empty($_SESSION['admin'])) {
    http_response_code(401);
    echo json_encode(['success' => false, 'error' => 'Unauthorised']);
    exit;
}

$id = (int)(isset($_POST['id']) ? $_POST['id'] : 0);
if (!$id) {
    echo json_encode(['success' => false, 'error' => 'Invalid topic ID']);
    exit;
}

// Fetch topic
$db   = getDB();
$stmt = $db->prepare('SELECT id, topic, level, chapters, front_matter, custom_toc, custom_instructions, phone, email, status FROM topics WHERE id = ?');
$stmt->execute([$id]);
$row = $stmt->fetch();

if (!$row) {
    echo json_encode(['success' => false, 'error' => 'Topic not found']);
    exit;
}
if ($row['status'] !== 'pending') {
    echo json_encode(['success' => false, 'error' => 'Topic has already been processed']);
    exit;
}

// POST to Flask /generate
// Parse front_matter string → array for Flask
$fm_string  = isset($row['front_matter']) ? trim($row['front_matter']) : '';
$fm_sections = ($fm_string !== '')
    ? array_values(array_filter(array_map('trim', explode(',', $fm_string))))
    : array();  // empty = abstract only

$body = array(
    'project_topic'         => $row['topic'],
    'research_level'        => $row['level'],
    'chapters'              => $row['chapters'],
    'front_matter_sections' => $fm_sections,
    'email'                 => $row['email'],
    'phone'                 => $row['phone'],
);
if (!empty($row['custom_toc'])) {
    $body['custom_toc'] = $row['custom_toc'];
}
if (!empty($row['custom_instructions'])) {
    $body['custom_instructions'] = $row['custom_instructions'];
}
$payload = json_encode($body);

$ch = curl_init(FLASK_API_URL);
curl_setopt_array($ch, [
    CURLOPT_POST           => true,
    CURLOPT_POSTFIELDS     => $payload,
    CURLOPT_RETURNTRANSFER => true,
    CURLOPT_HTTPHEADER     => ['Content-Type: application/json'],
    CURLOPT_TIMEOUT        => 30,
    CURLOPT_CONNECTTIMEOUT => 10,
]);
$response = curl_exec($ch);
$httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
$curlErr  = curl_error($ch);
curl_close($ch);

if ($curlErr) {
    echo json_encode(['success' => false, 'error' => "Could not reach API: $curlErr"]);
    exit;
}
if ($httpCode !== 200) {
    echo json_encode(['success' => false, 'error' => "API returned HTTP $httpCode: $response"]);
    exit;
}

$data  = json_decode($response, true);
$jobId = isset($data['job_id']) ? $data['job_id'] : null;

if (!$jobId) {
    echo json_encode(['success' => false, 'error' => 'API did not return a job_id']);
    exit;
}

// Update DB status → approved
$db->prepare(
    'UPDATE topics SET status = "approved", job_id = ?, approved_at = NOW() WHERE id = ?'
)->execute([$jobId, $id]);

echo json_encode(['success' => true, 'job_id' => $jobId, 'id' => $id]);

<?php
require_once __DIR__ . '/db.php';

header('Content-Type: application/json');

// Collect & sanitise
$topic        = trim(isset($_POST['topic'])        ? $_POST['topic']        : '');
$level        = trim(isset($_POST['level'])        ? $_POST['level']        : '');
$chapters     = trim(isset($_POST['chapters'])     ? $_POST['chapters']     : '1-5');
$front_matter = trim(isset($_POST['front_matter']) ? $_POST['front_matter'] : 'declaration,dedication,acknowledgements');
$phone        = trim(isset($_POST['phone'])        ? $_POST['phone']        : '');
$email        = trim(isset($_POST['email'])        ? $_POST['email']        : '');
$custom_toc          = trim(isset($_POST['custom_toc'])          ? $_POST['custom_toc']          : '');
$custom_instructions = trim(isset($_POST['custom_instructions']) ? $_POST['custom_instructions'] : '');

// Normalise empty chapters → default
if ($chapters === '') $chapters = '1-5';

// Validate and normalise front_matter
$allowed_fm = array('declaration', 'dedication', 'acknowledgements');
$fm_parts   = array();
foreach (explode(',', $front_matter) as $part) {
    $p = strtolower(trim($part));
    if (in_array($p, $allowed_fm, true)) $fm_parts[] = $p;
}
$front_matter = implode(',', $fm_parts); // may be empty string if all deselected

// Validate
$errors = [];

if ($topic === '')
    $errors[] = 'Research topic is required.';
elseif (mb_strlen($topic) > 600)
    $errors[] = 'Topic must be 600 characters or fewer.';

if (!in_array($level, array('undergraduate', 'postgraduate'), true))
    $errors[] = 'Invalid research level.';

// Validate chapters string: digits, commas, dashes, spaces only; must contain at least one valid 1-5 digit
if (!preg_match('/^[\d,\- ]+$/', $chapters))
    $errors[] = 'Invalid chapters format. Use e.g. "3", "3-5", or "1,3,5".';
else {
    // Parse and check at least one valid chapter number
    $valid = parseChapters($chapters);
    if (empty($valid))
        $errors[] = 'Please select at least one valid chapter (1–5).';
    else
        $chapters = buildChaptersString($valid); // normalise
}

if ($phone === '')
    $errors[] = 'Phone number is required.';

if (!filter_var($email, FILTER_VALIDATE_EMAIL))
    $errors[] = 'A valid email address is required.';

if ($errors) {
    echo json_encode(array('success' => false, 'errors' => $errors));
    exit;
}

// Insert into DB
try {
    $db   = getDB();
    $stmt = $db->prepare(
        'INSERT INTO topics (topic, level, chapters, front_matter, custom_toc, custom_instructions, phone, email)
         VALUES (:topic, :level, :chapters, :front_matter, :custom_toc, :custom_instructions, :phone, :email)'
    );
    $stmt->execute(array(
        ':topic'               => $topic,
        ':level'               => $level,
        ':chapters'            => $chapters,
        ':front_matter'        => ($front_matter        !== '') ? $front_matter        : null,
        ':custom_toc'          => ($custom_toc          !== '') ? $custom_toc          : null,
        ':custom_instructions' => ($custom_instructions !== '') ? $custom_instructions : null,
        ':phone'               => $phone,
        ':email'               => $email,
    ));
    echo json_encode(array('success' => true, 'id' => (int)$db->lastInsertId()));
} catch (PDOException $e) {
    echo json_encode(array('success' => false, 'errors' => array('Database error: ' . $e->getMessage())));
}

// ─────────────────────────────────────────────────────────
// Helper: parse chapters string into sorted array of 1-5
// ─────────────────────────────────────────────────────────
function parseChapters($str) {
    $result = array();
    $parts  = explode(',', $str);
    foreach ($parts as $part) {
        $part = trim($part);
        if ($part === '') continue;
        if (strpos($part, '-') !== false) {
            list($a, $b) = explode('-', $part, 2);
            $a = (int)trim($a); $b = (int)trim($b);
            for ($n = min($a,$b); $n <= max($a,$b); $n++) {
                if ($n >= 1 && $n <= 5) $result[] = $n;
            }
        } else {
            $n = (int)$part;
            if ($n >= 1 && $n <= 5) $result[] = $n;
        }
    }
    return array_values(array_unique($result));
}

// Helper: compress [1,2,3,5] → "1-3,5"
function buildChaptersString($nums) {
    sort($nums);
    $parts = array(); $i = 0;
    while ($i < count($nums)) {
        $start = $nums[$i]; $end = $nums[$i];
        while ($i + 1 < count($nums) && $nums[$i+1] === $nums[$i]+1) { $i++; $end = $nums[$i]; }
        $parts[] = ($start === $end) ? (string)$start : $start.'-'.$end;
        $i++;
    }
    return implode(',', $parts);
}

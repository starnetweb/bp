<?php
require_once __DIR__ . '/db.php';
require_once __DIR__ . '/config.php';

session_name(SESSION_NAME);
session_start();

// ── Login / Logout ─────────────────────────────────────────
$loginError = '';

if (isset($_POST['logout'])) {
    session_destroy();
    header('Location: admin.php');
    exit;
}

if (isset($_POST['password'])) {
    if ($_POST['password'] === ADMIN_PASSWORD) {
        $_SESSION['admin'] = true;
        header('Location: admin.php');
        exit;
    }
    $loginError = 'Incorrect password — please try again.';
}

$isLoggedIn = !empty($_SESSION['admin']);

// ── Fetch stats & rows (only when logged in) ──────────────
$stats = ['total' => 0, 'pending' => 0, 'approved' => 0, 'completed' => 0];
$rows  = [];

if ($isLoggedIn) {
    $db = getDB();

    $s = $db->query(
        "SELECT
            COUNT(*)                                              AS total,
            SUM(status = 'pending')                              AS pending,
            SUM(status IN ('approved','processing','completed')) AS approved,
            SUM(status = 'completed')                            AS completed
         FROM topics"
    )->fetch();
    $stats = $s ?: $stats;

    $rows = $db->query(
        "SELECT id, topic, level, chapters, front_matter, custom_toc, custom_instructions, phone, email, status, job_id, submitted_at
         FROM topics
         ORDER BY submitted_at DESC"
    )->fetchAll();
}

// ── Chapter label helper ─────────────────────────────────
function chLabel($chapters) {
    $s = trim((string)$chapters);
    // Normalised full set
    if ($s === '1-5' || $s === '1,2,3,4,5') return 'Complete (1-5)';
    // Single digit backward-compat integer "up to N"
    if (is_numeric($s) && (int)$s >= 1 && (int)$s <= 5) {
        $n = (int)$s;
        return $n === 1 ? 'Ch 1 only' : 'Ch 1-' . $n;
    }
    return 'Ch ' . $s;
}
?>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Admin — Research Portal</title>
<style>
/* ── Variables ───────────────────────────────────── */
:root{
  --bg:#0D1B2A;--surface:#112436;--surface2:#162E44;
  --border:#1F3A52;--accent:#3E8EE3;--accent-h:#5AA6F8;
  --success:#3CB97A;--error:#E05252;--warn:#F5A623;
  --pending:#F5A623;--approved:#3CB97A;
  --text:#DCE9F8;--muted:#6E90B0;--white:#fff;
  --radius:12px;
}
*{box-sizing:border-box;margin:0;padding:0}
body{
  background:var(--bg);color:var(--text);
  font-family:'Segoe UI',system-ui,sans-serif;
  min-height:100vh;display:flex;flex-direction:column;
}

/* ── Header ──────────────────────────────────────── */
header{
  background:linear-gradient(135deg,#0A1928,#132B42);
  border-bottom:1px solid var(--border);
  padding:16px 28px;
  display:flex;align-items:center;gap:14px;
}
.logo{
  width:40px;height:40px;background:var(--accent);
  border-radius:9px;display:flex;align-items:center;
  justify-content:center;font-size:1.2rem;flex-shrink:0;
}
.brand h1{font-size:1.15rem;font-weight:700;color:var(--white)}
.brand p{font-size:.75rem;color:var(--muted);margin-top:2px}
.header-actions{margin-left:auto;display:flex;align-items:center;gap:10px}
.btn-logout{
  font-size:.8rem;color:var(--muted);background:none;
  border:1px solid var(--border);border-radius:7px;
  padding:7px 14px;cursor:pointer;transition:all .2s;
}
.btn-logout:hover{color:var(--error);border-color:var(--error)}
.nav-link{
  font-size:.8rem;color:var(--muted);text-decoration:none;
  border:1px solid var(--border);border-radius:7px;
  padding:7px 14px;transition:all .2s;
}
.nav-link:hover{color:var(--accent);border-color:var(--accent)}

/* ── Login card ──────────────────────────────────── */
.login-wrap{
  flex:1;display:flex;align-items:center;
  justify-content:center;padding:36px 16px;
}
.login-card{
  background:var(--surface);border:1px solid var(--border);
  border-radius:var(--radius);padding:40px 44px;
  width:100%;max-width:400px;text-align:center;
  box-shadow:0 4px 28px rgba(0,0,0,.45);
}
.login-card .lock{
  font-size:2.8rem;margin-bottom:14px;
}
.login-card h2{
  font-size:1.3rem;font-weight:700;
  color:var(--white);margin-bottom:6px;
}
.login-card p{
  font-size:.84rem;color:var(--muted);margin-bottom:24px;
}
.login-card input[type=password]{
  width:100%;background:#0A1625;border:1.5px solid var(--border);
  border-radius:9px;color:var(--text);font-size:1rem;
  padding:12px 14px;outline:none;transition:border .2s;
  text-align:center;letter-spacing:.15em;
}
.login-card input:focus{border-color:var(--accent)}
.err-msg{
  background:#3D1B1B;border:1px solid var(--error);
  color:#F08080;border-radius:8px;
  padding:10px 14px;font-size:.83rem;
  margin-top:14px;display:<?= $loginError ? 'block' : 'none' ?>;
}
.btn-login{
  width:100%;margin-top:16px;padding:13px;
  background:var(--accent);color:#fff;border:none;
  border-radius:9px;font-size:1rem;font-weight:700;
  cursor:pointer;transition:background .2s;
}
.btn-login:hover{background:var(--accent-h)}

/* ── Stats bar ───────────────────────────────────── */
.stats-bar{
  display:grid;grid-template-columns:repeat(4,1fr);
  gap:14px;padding:20px 28px;
}
.stat-card{
  background:var(--surface);border:1px solid var(--border);
  border-radius:10px;padding:16px 20px;
  display:flex;align-items:center;gap:14px;
}
.stat-icon{
  width:44px;height:44px;border-radius:10px;
  display:flex;align-items:center;justify-content:center;
  font-size:1.3rem;flex-shrink:0;
}
.stat-icon.all{background:rgba(62,142,227,.12);color:var(--accent)}
.stat-icon.pend{background:rgba(245,166,35,.12);color:var(--warn)}
.stat-icon.appr{background:rgba(60,185,122,.12);color:var(--success)}
.stat-icon.done{background:rgba(90,166,248,.12);color:var(--accent-h)}
.stat-val{font-size:1.7rem;font-weight:800;color:var(--white)}
.stat-label{font-size:.73rem;color:var(--muted);margin-top:2px}

/* ── Table section ───────────────────────────────── */
.table-section{padding:0 28px 32px}
.table-header{
  display:flex;align-items:center;justify-content:space-between;
  margin-bottom:14px;flex-wrap:wrap;gap:10px;
}
.table-header h2{font-size:1rem;font-weight:700;color:var(--white)}
.search-box{
  background:#0A1625;border:1px solid var(--border);
  border-radius:8px;color:var(--text);font-size:.88rem;
  padding:8px 14px;outline:none;width:220px;transition:border .2s;
}
.search-box:focus{border-color:var(--accent)}

.table-wrap{
  background:var(--surface);border:1px solid var(--border);
  border-radius:var(--radius);overflow:hidden;overflow-x:auto;
}
table{width:100%;border-collapse:collapse;min-width:780px}
thead tr{background:#0D1F31}
th{
  padding:13px 16px;font-size:.72rem;font-weight:700;
  text-transform:uppercase;letter-spacing:.06em;color:var(--muted);
  text-align:left;white-space:nowrap;
}
tbody tr{
  border-top:1px solid var(--border);
  transition:background .15s;
}
tbody tr:hover{background:#132D45}
td{padding:13px 16px;font-size:.87rem;vertical-align:middle}
.topic-cell{
  max-width:240px;overflow:hidden;
  text-overflow:ellipsis;white-space:nowrap;
  color:var(--white);font-weight:600;
}
.level-badge{
  display:inline-block;padding:3px 10px;border-radius:20px;
  font-size:.7rem;font-weight:700;text-transform:uppercase;
  letter-spacing:.04em;
}
.level-ug{background:rgba(62,142,227,.15);color:var(--accent)}
.level-pg{background:rgba(90,166,248,.15);color:var(--accent-h)}

/* Status dots */
.status-dot{
  display:inline-flex;align-items:center;gap:6px;
  font-size:.8rem;font-weight:600;
}
.status-dot::before{
  content:'';width:9px;height:9px;border-radius:50%;flex-shrink:0;
}
.st-pending  .status-dot::before{background:var(--pending);box-shadow:0 0 6px var(--pending)}
.st-approved .status-dot::before{background:var(--approved);box-shadow:0 0 6px var(--approved)}
.st-processing .status-dot::before{background:var(--accent);box-shadow:0 0 6px var(--accent);animation:pulse 1.2s ease-in-out infinite}
.st-completed .status-dot::before{background:var(--success);box-shadow:0 0 6px var(--success)}
.st-failed   .status-dot::before{background:var(--error);box-shadow:0 0 6px var(--error)}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}

.st-pending   .status-dot{color:var(--pending)}
.st-approved  .status-dot{color:var(--approved)}
.st-processing .status-dot{color:var(--accent)}
.st-completed .status-dot{color:var(--success)}
.st-failed    .status-dot{color:var(--error)}

/* Action buttons */
.btn-approve{
  padding:7px 14px;background:var(--accent);color:#fff;
  border:none;border-radius:7px;font-size:.78rem;font-weight:700;
  cursor:pointer;transition:background .2s,transform .1s;
  white-space:nowrap;display:block;width:100%;margin-bottom:5px;
}
.btn-approve:hover{background:var(--accent-h)}
.btn-approve:active{transform:scale(.97)}
.btn-approve:disabled{
  background:#1F3A52;color:#3A6080;
  cursor:not-allowed;
}
.btn-download{
  padding:7px 14px;background:#1B3D2A;color:#3CB97A;
  border:1px solid #3CB97A;border-radius:7px;font-size:.78rem;font-weight:700;
  cursor:pointer;text-decoration:none;
  white-space:nowrap;display:block;width:100%;text-align:center;
  transition:background .2s,transform .1s;
}
.btn-download:hover{background:#243D2E;color:#7AE0A8;border-color:#7AE0A8}
.btn-download:active{transform:scale(.97)}
.action-cell{min-width:110px}
.job-id{
  font-size:.68rem;color:var(--muted);font-family:monospace;
  max-width:120px;overflow:hidden;text-overflow:ellipsis;
  white-space:nowrap;display:block;
}

/* empty state */
.empty{
  text-align:center;padding:52px 20px;
  color:var(--muted);font-size:.9rem;
}
.empty .icon{font-size:3rem;margin-bottom:12px}

/* ── Toast ───────────────────────────────────────── */
#toast{
  position:fixed;bottom:24px;right:24px;z-index:999;
  padding:13px 20px;border-radius:10px;font-size:.88rem;
  font-weight:600;max-width:360px;display:none;
  box-shadow:0 4px 24px rgba(0,0,0,.55);
  animation:slideup .3s ease;
}
@keyframes slideup{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:none}}
.toast-success{background:#1B3D2A;border:1px solid var(--success);color:#7AE0A8}
.toast-error  {background:#3D1B1B;border:1px solid var(--error);color:#F08080}

/* ── Mobile ──────────────────────────────────────── */
@media(max-width:700px){
  .stats-bar{grid-template-columns:1fr 1fr}
  .table-section,.stats-bar{padding-left:16px;padding-right:16px}
  header{padding:14px 16px}
}
@media(max-width:420px){
  .stats-bar{grid-template-columns:1fr}
  .login-card{padding:28px 20px}
}
</style>
</head>
<body>

<header>
  <div class="logo">🛡️</div>
  <div class="brand">
    <h1>Admin Dashboard</h1>
    <p>Research Portal Management</p>
  </div>
  <div class="header-actions">
    <a href="index.php" class="nav-link">📝 Submit Form</a>
    <?php if ($isLoggedIn): ?>
    <form method="POST" style="margin:0">
      <button name="logout" class="btn-logout">🚪 Logout</button>
    </form>
    <?php endif; ?>
  </div>
</header>

<?php if (!$isLoggedIn): ?>
<!-- ════ LOGIN SCREEN ════ -->
<div class="login-wrap">
  <div class="login-card">
    <div class="lock">🔐</div>
    <h2>Admin Access</h2>
    <p>Enter your admin password to continue</p>
    <form method="POST">
      <input type="password" name="password"
             placeholder="••••••••••" autocomplete="current-password" autofocus/>
      <?php if ($loginError): ?>
      <div class="err-msg"><?= htmlspecialchars($loginError) ?></div>
      <?php endif; ?>
      <button type="submit" class="btn-login">🔓 Sign In</button>
    </form>
  </div>
</div>

<?php else: ?>
<!-- ════ DASHBOARD ════ -->

<!-- Stats bar -->
<div class="stats-bar">
  <div class="stat-card">
    <div class="stat-icon all">📋</div>
    <div>
      <div class="stat-val"><?= (int)$stats['total'] ?></div>
      <div class="stat-label">Total Submitted</div>
    </div>
  </div>
  <div class="stat-card">
    <div class="stat-icon pend">⏳</div>
    <div>
      <div class="stat-val"><?= (int)$stats['pending'] ?></div>
      <div class="stat-label">Pending Approval</div>
    </div>
  </div>
  <div class="stat-card">
    <div class="stat-icon appr">✅</div>
    <div>
      <div class="stat-val"><?= (int)$stats['approved'] ?></div>
      <div class="stat-label">Approved / Processing</div>
    </div>
  </div>
  <div class="stat-card">
    <div class="stat-icon done">📄</div>
    <div>
      <div class="stat-val"><?= (int)$stats['completed'] ?></div>
      <div class="stat-label">Completed</div>
    </div>
  </div>
</div>

<!-- Table -->
<div class="table-section">
  <div class="table-header">
    <h2>📋 Submitted Topics</h2>
    <input class="search-box" type="text" id="search"
           placeholder="🔍 Search topics…" oninput="filterTable()"/>
  </div>

  <div class="table-wrap">
    <?php if (empty($rows)): ?>
    <div class="empty">
      <div class="icon">📭</div>
      <p>No topics have been submitted yet.<br>
         <a href="index.php" style="color:var(--accent)">Submit the first one →</a></p>
    </div>
    <?php else: ?>
    <table id="topics-table">
      <thead>
        <tr>
          <th>#</th>
          <th>Topic</th>
          <th>Level</th>
          <th>Chapters</th>
          <th>Phone</th>
          <th>Email</th>
          <th>Status</th>
          <th>Submitted</th>
          <th>Action</th>
        </tr>
      </thead>
      <tbody>
        <?php foreach ($rows as $r):
          $stClass = 'st-' . htmlspecialchars($r['status']);
          $stLabel = ucfirst($r['status']);
          $isPending = ($r['status'] === 'pending');
          $submitted = date('d M Y, H:i', strtotime($r['submitted_at']));
        ?>
        <tr class="<?= $stClass ?>" data-search="<?= strtolower(htmlspecialchars($r['topic'].' '.$r['email'].' '.$r['phone'])) ?>">
          <td style="color:var(--muted)"><?= (int)$r['id'] ?></td>
          <td>
            <div class="topic-cell" title="<?= htmlspecialchars($r['topic']) ?>">
              <?= htmlspecialchars($r['topic']) ?>
            </div>
          </td>
          <td>
            <span class="level-badge <?= $r['level'] === 'undergraduate' ? 'level-ug' : 'level-pg' ?>">
              <?= $r['level'] === 'undergraduate' ? 'UG' : 'PG' ?>
            </span>
          </td>
          <td style="color:var(--muted);font-size:.82rem">
            <?= htmlspecialchars(chLabel($r['chapters'])) ?>
            <?php
              // Front matter badges
              $fm_all  = array('declaration','dedication','acknowledgements');
              $fm_have = ($r['front_matter'] !== null && $r['front_matter'] !== '')
                         ? array_map('trim', explode(',', $r['front_matter']))
                         : array();
              $fm_miss = array_diff($fm_all, $fm_have);
              foreach ($fm_miss as $ms):
            ?>
              <span style="color:var(--error);font-size:.65rem;display:block">✗ <?= ucfirst($ms) ?></span>
            <?php endforeach; ?>
            <?php if (!empty($r['custom_toc'])): ?>
              <span title="Custom TOC provided" style="color:var(--accent);font-size:.68rem;display:block">📋 Custom TOC</span>
            <?php endif; ?>
            <?php if (!empty($r['custom_instructions'])): ?>
              <span title="<?= htmlspecialchars(mb_substr($r['custom_instructions'], 0, 120), ENT_QUOTES, 'UTF-8') ?>" style="color:var(--warn);font-size:.68rem;display:block">📝 Custom Instructions</span>
            <?php endif; ?>
          </td>
          <td style="color:var(--muted);font-size:.83rem"><?= htmlspecialchars($r['phone']) ?></td>
          <td style="color:var(--muted);font-size:.83rem"><?= htmlspecialchars($r['email']) ?></td>
          <td>
            <div class="status-dot"><?= $stLabel ?></div>
            <?php if ($r['job_id']): ?>
            <span class="job-id" title="Flask Job ID: <?= htmlspecialchars($r['job_id']) ?>">
              <?= htmlspecialchars(substr($r['job_id'], 0, 8)) ?>…
            </span>
            <?php endif; ?>
          </td>
          <td style="color:var(--muted);font-size:.8rem;white-space:nowrap"><?= $submitted ?></td>
          <td class="action-cell">
            <?php if ($isPending): ?>
            <button class="btn-approve"
                    data-id="<?= (int)$r['id'] ?>"
                    data-topic="<?= htmlspecialchars($r['topic'], ENT_QUOTES, 'UTF-8') ?>"
                    onclick="approve(this)">
              ✅ Approve
            </button>
            <?php else: ?>
            <button class="btn-approve" disabled style="margin-bottom:<?= $r['job_id'] ? '5px' : '0' ?>">—</button>
            <?php endif; ?>
            <?php if ($r['job_id']): ?>
            <a class="btn-download"
               href="download.php?id=<?= (int)$r['id'] ?>"
               title="Download completed document">
              📥 Download
            </a>
            <?php endif; ?>
          </td>
        </tr>
        <?php endforeach; ?>
      </tbody>
    </table>
    <?php endif; ?>
  </div>
</div>

<div id="toast"></div>

<script>
// ── Table search ───────────────────────────────────────────
function filterTable() {
  const q = document.getElementById('search').value.toLowerCase();
  document.querySelectorAll('#topics-table tbody tr').forEach(row => {
    row.style.display = row.dataset.search.includes(q) ? '' : 'none';
  });
}

// ── Toast helper ───────────────────────────────────────────
function toast(msg, type = 'success') {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'toast-' + type;
  t.style.display = 'block';
  setTimeout(() => t.style.display = 'none', 4500);
}

// ── Approve handler ────────────────────────────────────────
async function approve(btn) {
  var id    = btn.dataset.id;
  var topic = btn.dataset.topic;
  var shortTopic = topic.length > 55 ? topic.slice(0, 55) + '…' : topic;
  if (!confirm(`Approve and send to Flask API?\n\n"${shortTopic}"`)) return;

  btn.disabled    = true;
  btn.textContent = '⏳ Sending…';

  const fd = new FormData();
  fd.append('id', id);

  try {
    const res  = await fetch('approve.php', { method: 'POST', body: fd });
    const data = await res.json();

    if (data.success) {
      // Update the row status to "approved" without page reload
      const row = btn.closest('tr');
      row.className = 'st-approved';
      row.querySelector('.status-dot').textContent = 'Approved';

      const jobSpan = document.createElement('span');
      jobSpan.className = 'job-id';
      jobSpan.title = 'Flask Job ID: ' + data.job_id;
      jobSpan.textContent = data.job_id.slice(0, 8) + '…';
      row.querySelector('td:nth-child(7)').appendChild(jobSpan);

      // Disable approve btn and inject Download button
      btn.textContent = '—';
      btn.style.marginBottom = '5px';
      const dlLink = document.createElement('a');
      dlLink.className = 'btn-download';
      dlLink.href = 'download.php?id=' + data.id;
      dlLink.title = 'Download completed document';
      dlLink.textContent = '📥 Download';
      btn.parentNode.appendChild(dlLink);

      toast('✅ Approved! Job ID: ' + data.job_id.slice(0, 8) + '…');

      // Update stat badges live
      const pendEl = document.querySelector('.stat-icon.pend + div .stat-val');
      if (pendEl) pendEl.textContent = Math.max(0, parseInt(pendEl.textContent) - 1);
      const apprEl = document.querySelector('.stat-icon.appr + div .stat-val');
      if (apprEl) apprEl.textContent = parseInt(apprEl.textContent) + 1;
    } else {
      btn.disabled    = false;
      btn.textContent = '✅ Approve';
      toast('❌ ' + (data.error || 'Approval failed'), 'error');
    }
  } catch (err) {
    btn.disabled    = false;
    btn.textContent = '✅ Approve';
    toast('❌ Network error — please try again.', 'error');
  }
}
</script>

<?php endif; ?>

<footer style="text-align:center;padding:18px;color:var(--muted);font-size:.73rem;border-top:1px solid var(--border);margin-top:auto">
  &copy; <?= date('Y') ?> Research Portal · Admin Dashboard
</footer>
</body>
</html>

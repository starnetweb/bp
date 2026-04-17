<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Submit Research Topic — Research Portal</title>
<style>
:root{
  --bg:#0D1B2A;--surface:#112436;--surface2:#162E44;
  --border:#1F3A52;--accent:#3E8EE3;--accent-h:#5AA6F8;
  --success:#3CB97A;--error:#E05252;--warn:#F5A623;
  --text:#DCE9F8;--muted:#6E90B0;--white:#fff;
  --radius:12px;--shadow:0 4px 24px rgba(0,0,0,.4);
}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--text);font-family:'Segoe UI',system-ui,sans-serif;min-height:100vh;display:flex;flex-direction:column}

header{background:linear-gradient(135deg,#0A1928,#132B42);border-bottom:1px solid var(--border);padding:18px 24px;display:flex;align-items:center;gap:14px}
.logo{width:42px;height:42px;background:var(--accent);border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:1.3rem;flex-shrink:0}
.brand h1{font-size:1.2rem;font-weight:700;color:var(--white)}
.brand p{font-size:.78rem;color:var(--muted);margin-top:2px}
.nav-link{margin-left:auto;font-size:.82rem;color:var(--muted);text-decoration:none;padding:7px 14px;border:1px solid var(--border);border-radius:7px;transition:all .2s;white-space:nowrap}
.nav-link:hover{color:var(--accent);border-color:var(--accent)}

main{flex:1;display:flex;align-items:flex-start;justify-content:center;padding:36px 16px}
.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:36px 40px;width:100%;max-width:620px;box-shadow:var(--shadow)}
.card-head{margin-bottom:28px}
.card-head h2{font-size:1.35rem;font-weight:700;color:var(--white)}
.card-head p{font-size:.85rem;color:var(--muted);margin-top:5px}

.form-group{margin-bottom:20px}
.form-group label{display:block;font-size:.75rem;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:7px}
.form-group input[type=text],.form-group input[type=email],.form-group input[type=tel]{width:100%;background:#0A1625;border:1.5px solid var(--border);border-radius:9px;color:var(--text);font-size:.97rem;padding:12px 14px;outline:none;transition:border .2s,box-shadow .2s}
.form-group input:focus{border-color:var(--accent);box-shadow:0 0 0 3px rgba(62,142,227,.15)}
.form-group input::placeholder{color:#3A5A7A}
.row-2{display:grid;grid-template-columns:1fr 1fr;gap:16px}

/* Level */
.level-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.level-card{background:#0A1625;border:2px solid var(--border);border-radius:10px;padding:16px 14px;cursor:pointer;transition:border .2s,background .2s;text-align:center;user-select:none}
.level-card:hover{border-color:var(--accent);background:#0F2035}
.level-card.active{border-color:var(--accent);background:#0F2035;box-shadow:0 0 0 3px rgba(62,142,227,.12)}
.level-icon{font-size:1.5rem;margin-bottom:7px}
.level-title{font-size:.93rem;font-weight:700;color:var(--white)}
.level-desc{font-size:.73rem;color:var(--muted);margin-top:4px;line-height:1.45}
input[name=level]{display:none}

/* Chapter toggles */
.chapter-grid{display:flex;gap:8px;flex-wrap:wrap}
.ch-btn{flex:1;min-width:56px;padding:10px 6px;text-align:center;background:#0A1625;border:1.5px solid var(--border);border-radius:8px;cursor:pointer;font-size:.88rem;font-weight:700;color:var(--muted);transition:all .2s;user-select:none}
.ch-btn:hover{border-color:var(--accent);color:var(--accent)}
.ch-btn.active{background:var(--accent);border-color:var(--accent);color:#fff;box-shadow:0 2px 10px rgba(62,142,227,.3)}
.ch-label{font-size:.62rem;display:block;margin-top:3px;opacity:.8;font-weight:400}
.ch-quick{display:flex;gap:6px;margin-top:8px;flex-wrap:wrap}
.ch-quick button{background:none;border:1px solid var(--border);border-radius:5px;color:var(--muted);font-size:.72rem;padding:4px 10px;cursor:pointer;transition:all .2s}
.ch-quick button:hover{border-color:var(--accent);color:var(--accent)}
.ch-summary{font-size:.78rem;color:var(--accent);margin-top:6px;min-height:1.2em}
input[name=chapters]{display:none}

/* Front matter toggles */
.fm-checks{display:flex;gap:10px;flex-wrap:wrap;margin-top:6px}
.fm-check{display:flex;align-items:center;gap:7px;background:#0A1625;border:1.5px solid var(--border);border-radius:8px;padding:9px 14px;cursor:pointer;user-select:none;transition:border .2s,background .2s}
.fm-check:hover{border-color:var(--accent)}
.fm-check.active{border-color:var(--accent);background:#0F2035}
.fm-check-icon{font-size:1rem}
.fm-check-lbl{font-size:.82rem;font-weight:600;color:var(--text)}
.fm-note{font-size:.73rem;color:var(--muted);margin-top:7px}

/* Custom TOC */
.toc-toggle{display:flex;align-items:center;gap:8px;cursor:pointer;user-select:none;margin-top:4px}
.toc-toggle input[type=checkbox]{width:16px;height:16px;accent-color:var(--accent);cursor:pointer;flex-shrink:0}
.toc-toggle span{font-size:.83rem;color:var(--muted)}
#custom-toc-wrap{display:none;margin-top:10px}
#custom-toc{width:100%;background:#0A1625;border:1.5px solid var(--border);border-radius:9px;color:var(--text);font-size:.82rem;font-family:Consolas,monospace;padding:10px 12px;outline:none;resize:vertical;min-height:140px;transition:border .2s;line-height:1.6}
#custom-toc:focus{border-color:var(--accent)}
.toc-hint{font-size:.72rem;color:var(--muted);margin-top:5px;line-height:1.5}

/* Custom instructions */
.ci-toggle{display:flex;align-items:center;gap:8px;cursor:pointer;user-select:none;margin-top:4px}
.ci-toggle input[type=checkbox]{width:16px;height:16px;accent-color:var(--accent);cursor:pointer;flex-shrink:0}
.ci-toggle span{font-size:.83rem;color:var(--muted)}
#ci-wrap{display:none;margin-top:10px}
#custom-instructions-field{width:100%;background:#0A1625;border:1.5px solid var(--border);border-radius:9px;color:var(--text);font-size:.87rem;padding:11px 14px;outline:none;resize:vertical;min-height:110px;transition:border .2s;line-height:1.65}
#custom-instructions-field:focus{border-color:var(--accent)}
#custom-instructions-field::placeholder{color:#3A5A7A}
.ci-hint{font-size:.72rem;color:var(--muted);margin-top:5px;line-height:1.5}

.divider{border:none;border-top:1px solid var(--border);margin:24px 0}

.btn-submit{width:100%;padding:14px;background:var(--accent);color:#fff;border:none;border-radius:10px;font-size:1rem;font-weight:700;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:9px;transition:background .2s,transform .1s}
.btn-submit:hover{background:var(--accent-h)}
.btn-submit:active{transform:scale(.98)}
.btn-submit:disabled{background:#1F3A52;color:#4A6A8A;cursor:not-allowed}

#toast{position:fixed;top:20px;right:20px;z-index:999;padding:14px 20px;border-radius:10px;font-size:.9rem;font-weight:600;max-width:360px;display:none;box-shadow:0 4px 20px rgba(0,0,0,.5);animation:slidein .3s ease}
@keyframes slidein{from{opacity:0;transform:translateY(-12px)}to{opacity:1;transform:none}}
.toast-success{background:#1B3D2A;border:1px solid #3CB97A;color:#7AE0A8}
.toast-error{background:#3D1B1B;border:1px solid #E05252;color:#F08080}

#success-panel{display:none;text-align:center;padding:12px 0}
.check-circle{width:72px;height:72px;background:rgba(60,185,122,.12);border:2px solid var(--success);border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:2rem;margin:0 auto 16px}
#success-panel h3{color:var(--success);font-size:1.2rem;margin-bottom:8px}
#success-panel p{color:var(--muted);font-size:.87rem;line-height:1.6}
.btn-another{display:inline-block;margin-top:18px;padding:10px 24px;background:var(--surface2);border:1px solid var(--border);color:var(--text);border-radius:8px;cursor:pointer;font-size:.9rem;font-weight:600;transition:all .2s}
.btn-another:hover{border-color:var(--accent);color:var(--accent)}

footer{text-align:center;padding:18px;color:var(--muted);font-size:.75rem;border-top:1px solid var(--border)}

@media(max-width:540px){
  .card{padding:26px 18px}
  .row-2,.level-grid{grid-template-columns:1fr}
}
</style>
</head>
<body>

<header>
  <div class="logo">📝</div>
  <div class="brand">
    <h1>Research Portal</h1>
    <p>Academic Project Submission System</p>
  </div>
  <a href="admin.php" class="nav-link">🔐 Admin</a>
</header>

<main>
  <div class="card">
    <div class="card-head">
      <h2>Submit a Research Topic</h2>
      <p>Complete the form below to submit a paid research topic for processing.</p>
    </div>

    <div id="form-wrap">
      <form id="submit-form" novalidate>

        <div class="form-group">
          <label>Research Topic <span style="color:var(--error)">*</span></label>
          <input type="text" name="topic" id="topic"
                 placeholder="e.g. The Impact of Digital Finance on SME Growth in West Africa"
                 autocomplete="off" required/>
        </div>

        <div class="form-group">
          <label>Research Level <span style="color:var(--error)">*</span></label>
          <div class="level-grid">
            <div class="level-card active" data-val="undergraduate" onclick="setLevel(this)">
              <div class="level-icon">🎓</div>
              <div class="level-title">Undergraduate</div>
              <div class="level-desc">Structured analysis, solid methodology</div>
            </div>
            <div class="level-card" data-val="postgraduate" onclick="setLevel(this)">
              <div class="level-icon">🏛️</div>
              <div class="level-title">Postgraduate</div>
              <div class="level-desc">Critical theory, epistemological depth</div>
            </div>
          </div>
          <input type="hidden" name="level" id="level-input" value="undergraduate"/>
        </div>

        <div class="form-group">
          <label>Chapters to Generate <span style="color:var(--error)">*</span></label>
          <div class="chapter-grid">
            <div class="ch-btn active" data-ch="1" onclick="toggleChapter(this)">1<span class="ch-label">Intro</span></div>
            <div class="ch-btn active" data-ch="2" onclick="toggleChapter(this)">2<span class="ch-label">Lit. Rev.</span></div>
            <div class="ch-btn active" data-ch="3" onclick="toggleChapter(this)">3<span class="ch-label">Methodology</span></div>
            <div class="ch-btn active" data-ch="4" onclick="toggleChapter(this)">4<span class="ch-label">Results</span></div>
            <div class="ch-btn active" data-ch="5" onclick="toggleChapter(this)">5<span class="ch-label">Conclusions</span></div>
          </div>
          <div class="ch-quick">
            <button type="button" onclick="setChapters('all')">All</button>
            <button type="button" onclick="setChapters('1')">Ch 1 only</button>
            <button type="button" onclick="setChapters('1-3')">Ch 1–3</button>
            <button type="button" onclick="setChapters('3-5')">Ch 3–5</button>
            <button type="button" onclick="setChapters('none')">Clear</button>
          </div>
          <div class="ch-summary" id="ch-summary">Chapters 1-5 (complete)</div>
          <input type="hidden" name="chapters" id="chapters-input" value="1-5"/>
        </div>

        <div class="form-group">
          <label>Front Matter Sections</label>
          <div class="fm-checks">
            <div class="fm-check active" data-fm="declaration" onclick="toggleFm(this)">
              <span class="fm-check-icon">📜</span>
              <span class="fm-check-lbl">Declaration</span>
            </div>
            <div class="fm-check active" data-fm="dedication" onclick="toggleFm(this)">
              <span class="fm-check-icon">❤️</span>
              <span class="fm-check-lbl">Dedication</span>
            </div>
            <div class="fm-check active" data-fm="acknowledgements" onclick="toggleFm(this)">
              <span class="fm-check-icon">🙏</span>
              <span class="fm-check-lbl">Acknowledgements</span>
            </div>
          </div>
          <div class="fm-note">ℹ️ Abstract is always included. Click a section to include or exclude it.</div>
          <input type="hidden" name="front_matter" id="front-matter-input" value="declaration,dedication,acknowledgements"/>
        </div>

        <div class="form-group">
          <label>Table of Contents</label>
          <label class="toc-toggle">
            <input type="checkbox" id="toc-checkbox" onchange="toggleToc()"/>
            <span>Provide a custom table of contents</span>
          </label>
          <div id="custom-toc-wrap">
            <textarea id="custom-toc" name="custom_toc"
                      placeholder="Enter your TOC, one item per line. Example:&#10;CHAPTER ONE: INTRODUCTION&#10;  1.1  Background of the Study&#10;  1.2  Statement of the Problem&#10;&#10;CHAPTER TWO: LITERATURE REVIEW&#10;  2.1  Conceptual Framework&#10;  ..."></textarea>
            <div class="toc-hint">💡 Each line = one TOC entry. Leave blank lines between chapters for spacing.</div>
          </div>
        </div>

        <div class="form-group">
          <label>Custom Instructions</label>
          <label class="ci-toggle">
            <input type="checkbox" id="ci-checkbox" onchange="toggleCi()"/>
            <span>Add custom instructions for document generation</span>
          </label>
          <div id="ci-wrap">
            <textarea id="custom-instructions-field" name="custom_instructions"
              placeholder="Enter any specific instructions the AI should follow.&#10;&#10;Examples:&#10;• Use Nigeria as the primary case study&#10;• Cite specific scholars or theoretical frameworks&#10;• Focus on quantitative methods only&#10;• Follow a particular institution's formatting guidelines"></textarea>
            <div class="ci-hint">💡 These instructions will be applied to every chapter and front matter section generated.</div>
          </div>
        </div>

        <hr class="divider"/>

        <div class="row-2">
          <div class="form-group" style="margin-bottom:0">
            <label>Phone Number <span style="color:var(--error)">*</span></label>
            <input type="tel" name="phone" id="phone" placeholder="+234 800 000 0000" required/>
          </div>
          <div class="form-group" style="margin-bottom:0">
            <label>Email Address <span style="color:var(--error)">*</span></label>
            <input type="email" name="email" id="email" placeholder="client@example.com" required/>
          </div>
        </div>

        <div style="margin-top:28px">
          <button type="submit" class="btn-submit" id="sub-btn">
            <span id="btn-text">📤 Submit Research Topic</span>
          </button>
        </div>

      </form>
    </div>

    <div id="success-panel">
      <div class="check-circle">✓</div>
      <h3>Topic Submitted Successfully!</h3>
      <p>Your research topic has been received and is<br/>
         <strong style="color:var(--text)">pending admin approval</strong>.<br/>
         You will be contacted once the document is ready.</p>
      <div class="btn-another" onclick="resetForm()">⬅ Submit Another Topic</div>
    </div>
  </div>
</main>

<div id="toast"></div>

<footer>&copy; <?= date('Y') ?> Research Portal · Academic Project Submission System</footer>

<script>
// ── Level ─────────────────────────────────────────────────
function setLevel(el) {
  document.querySelectorAll('.level-card').forEach(function(c){ c.classList.remove('active'); });
  el.classList.add('active');
  document.getElementById('level-input').value = el.dataset.val;
}

// ── Chapters ──────────────────────────────────────────────
function getSelected() {
  return Array.prototype.slice.call(document.querySelectorAll('.ch-btn.active'))
    .map(function(b){ return parseInt(b.dataset.ch); })
    .sort(function(a,b){ return a-b; });
}

function buildChaptersString(sel) {
  if (!sel.length) return '';
  // Compress into ranges
  var parts = [], i = 0;
  while (i < sel.length) {
    var start = sel[i], end = sel[i];
    while (i + 1 < sel.length && sel[i+1] === sel[i]+1) { i++; end = sel[i]; }
    parts.push(start === end ? String(start) : start + '-' + end);
    i++;
  }
  return parts.join(',');
}

function updateChSummary() {
  var sel = getSelected();
  var el  = document.getElementById('ch-summary');
  var inp = document.getElementById('chapters-input');
  if (!sel.length) {
    el.textContent = 'No chapters selected'; el.style.color = 'var(--error)';
    inp.value = ''; return;
  }
  inp.value = buildChaptersString(sel);
  var all = [1,2,3,4,5];
  var label;
  if (JSON.stringify(sel) === JSON.stringify(all)) { label = 'Chapters 1-5 (complete)'; }
  else if (sel.length === 1) { label = 'Chapter ' + sel[0] + ' only'; }
  else {
    var isRange = sel.every(function(v,i){ return i===0 || v===sel[i-1]+1; });
    label = 'Chapters ' + (isRange ? sel[0]+'-'+sel[sel.length-1] : sel.join(', '));
  }
  el.textContent = label; el.style.color = 'var(--accent)';
}

function toggleChapter(el) { el.classList.toggle('active'); updateChSummary(); }

function setChapters(preset) {
  document.querySelectorAll('.ch-btn').forEach(function(b) {
    var n = parseInt(b.dataset.ch);
    if      (preset === 'all')  b.classList.add('active');
    else if (preset === 'none') b.classList.remove('active');
    else if (preset === '1')    b.classList.toggle('active', n === 1);
    else if (preset === '1-3')  b.classList.toggle('active', n <= 3);
    else if (preset === '3-5')  b.classList.toggle('active', n >= 3);
  });
  updateChSummary();
}

// ── Front matter toggles ─────────────────────────────────
function toggleFm(el) {
  el.classList.toggle('active');
  var vals = [];
  document.querySelectorAll('.fm-check.active').forEach(function(e){ vals.push(e.dataset.fm); });
  document.getElementById('front-matter-input').value = vals.join(',');
}

// ── Custom TOC ────────────────────────────────────────────
function toggleToc() {
  var show = document.getElementById('toc-checkbox').checked;
  document.getElementById('custom-toc-wrap').style.display = show ? 'block' : 'none';
}

// ── Custom instructions ───────────────────────────────────
function toggleCi() {
  var show = document.getElementById('ci-checkbox').checked;
  document.getElementById('ci-wrap').style.display = show ? 'block' : 'none';
}

// ── Toast ─────────────────────────────────────────────────
function toast(msg, type) {
  type = type || 'error';
  var t = document.getElementById('toast');
  t.textContent = msg; t.className = 'toast-' + type;
  t.style.display = 'block';
  setTimeout(function(){ t.style.display = 'none'; }, 4000);
}

// ── Submit ────────────────────────────────────────────────
document.getElementById('submit-form').addEventListener('submit', function(e) {
  e.preventDefault();
  var btn   = document.getElementById('sub-btn');
  var text  = document.getElementById('btn-text');
  var topic = document.getElementById('topic').value.trim();
  var phone = document.getElementById('phone').value.trim();
  var email = document.getElementById('email').value.trim();
  var chs   = document.getElementById('chapters-input').value;

  if (!topic)       { toast('Please enter the research topic.'); return; }
  if (!chs)         { toast('Please select at least one chapter.'); return; }
  if (!phone)       { toast('Please enter a phone number.'); return; }
  if (!email || email.indexOf('@') === -1) { toast('Please enter a valid email.'); return; }

  btn.disabled = true; text.textContent = '⏳ Submitting…';

  var fd = new FormData(document.getElementById('submit-form'));
  // If custom TOC box unchecked, clear the value so it isn't sent
  if (!document.getElementById('toc-checkbox').checked) {
    fd.set('custom_toc', '');
  }
  // Same for custom instructions
  if (!document.getElementById('ci-checkbox').checked) {
    fd.set('custom_instructions', '');
  }

  fetch('submit.php', { method: 'POST', body: fd })
    .then(function(r){ return r.json(); })
    .then(function(data) {
      if (data.success) {
        document.getElementById('form-wrap').style.display    = 'none';
        document.getElementById('success-panel').style.display = 'block';
      } else {
        toast((data.errors || ['Submission failed']).join('\n'));
        btn.disabled = false; text.textContent = '📤 Submit Research Topic';
      }
    })
    .catch(function() {
      toast('Network error — please try again.');
      btn.disabled = false; text.textContent = '📤 Submit Research Topic';
    });
});

// ── Reset ─────────────────────────────────────────────────
function resetForm() {
  document.getElementById('submit-form').reset();
  document.getElementById('level-input').value = 'undergraduate';
  document.querySelectorAll('.level-card').forEach(function(c,i){ c.classList.toggle('active', i===0); });
  document.querySelectorAll('.fm-check').forEach(function(c){ c.classList.add('active'); });
  document.getElementById('front-matter-input').value = 'declaration,dedication,acknowledgements';
  setChapters('all');
  document.getElementById('toc-checkbox').checked = false;
  document.getElementById('custom-toc-wrap').style.display = 'none';
  document.getElementById('ci-checkbox').checked = false;
  document.getElementById('ci-wrap').style.display = 'none';
  document.getElementById('sub-btn').disabled = false;
  document.getElementById('btn-text').textContent = '📤 Submit Research Topic';
  document.getElementById('form-wrap').style.display    = 'block';
  document.getElementById('success-panel').style.display = 'none';
}
</script>
</body>
</html>

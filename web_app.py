"""
Academic Research Writeup Agent — Web Interface
Run:  python web_app.py
Open: http://localhost:5000

POST /generate  JSON: { "project_topic": "...", "research_level": "undergraduate"|"postgraduate" }
"""

import os
import sys
import re
import uuid
import queue
import threading
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base       import MIMEBase
from email.mime.text       import MIMEText
from email                 import encoders

from flask import (
    Flask, render_template_string, request,
    jsonify, Response, send_file, abort
)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config
import research_agent

app = Flask(__name__)
app.secret_key = os.urandom(24)

JOBS: dict[str, dict] = {}
OUTPUT_DIR = (
    "/app/downloads"
    if os.path.exists("/app")
    else os.path.join(os.path.dirname(os.path.abspath(__file__)), "downloads")
)
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ─────────────────────────────────────────────────────────
#  EMAIL HELPER
# ─────────────────────────────────────────────────────────

def send_email(file_path: str, filename: str, topic: str,
               research_level: str, log_fn=None,
               extra_recipients: list | None = None):
    """Send the finished .docx to config.RECIPIENT_EMAILS + any extra_recipients."""
    def _log(msg):
        if log_fn:
            log_fn(msg)

    recipients = list(config.RECIPIENT_EMAILS)
    if extra_recipients:
        for em in extra_recipients:
            if em and em not in recipients:
                recipients.append(em)

    if not recipients:
        _log("  ⚠  No recipient emails configured — skipping email.")
        return

    _log(f"  Sending email to: {', '.join(recipients)}")

    level_label = research_agent.LEVEL_PROFILES[research_level]["label"]
    subject     = f"Research Document Ready: {topic[:60]}"
    body        = (
        f"Hello,\n\n"
        f"Your {level_label} research document has been generated.\n\n"
        f"Topic          : {topic}\n"
        f"Research Level : {level_label}\n"
        f"Document       : {filename}\n\n"
        f"The document is attached to this email.\n\n"
        f"Best regards,\nAcademic Research Agent"
    )

    msg = MIMEMultipart()
    msg["From"]    = config.SMTP_USER
    msg["To"]      = ", ".join(recipients)
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    # Attach the .docx
    with open(file_path, "rb") as f:
        part = MIMEBase("application",
                        "vnd.openxmlformats-officedocument.wordprocessingml.document")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
    msg.attach(part)

    try:
        with smtplib.SMTP(config.SMTP_HOST, config.SMTP_PORT, timeout=30) as server:
            server.ehlo()
            server.starttls()
            server.login(config.SMTP_USER, config.SMTP_PASSWORD)
            server.sendmail(config.SMTP_USER, recipients, msg.as_string())
        _log(f"  ✓ Email sent successfully to {len(recipients)} recipient(s)")
    except Exception as exc:
        _log(f"  ⚠  Email failed: {exc}")


# ─────────────────────────────────────────────────────────
#  HTML TEMPLATE
# ─────────────────────────────────────────────────────────
HTML = r"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Academic Research Agent</title>
<style>
:root{
  --bg:#0F1923;--card:#172130;--border:#253347;
  --accent:#4A90D9;--accent2:#5BA3F5;
  --success:#4CAF50;--warn:#FFA726;--error:#EF5350;
  --text:#E8F0FE;--muted:#7A9BBF;
}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--text);font-family:'Segoe UI',system-ui,sans-serif;min-height:100vh}

header{background:var(--card);border-bottom:1px solid var(--border);padding:22px 0;text-align:center}
header h1{font-size:1.7rem;font-weight:700;color:#fff}
header p{color:var(--muted);font-size:.9rem;margin-top:4px}

main{max-width:780px;margin:32px auto;padding:0 16px}

.card{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:28px 32px;margin-bottom:20px}
.card h2{font-size:1rem;font-weight:700;color:var(--accent);margin-bottom:16px;letter-spacing:.04em;text-transform:uppercase}

label{display:block;font-size:.8rem;font-weight:700;color:var(--muted);margin-bottom:6px;text-transform:uppercase;letter-spacing:.05em}
input[type=text],input[type=email],input[type=tel]{width:100%;background:#0F1923;border:1px solid var(--border);border-radius:8px;color:var(--text);font-size:1rem;padding:12px 14px;outline:none;transition:border .2s}
input:focus{border-color:var(--accent)}
.form-group{margin-bottom:20px}
.row-2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.divider{border:none;border-top:1px solid var(--border);margin:4px 0 20px}

/* Chapter selector */
.chapter-grid{display:flex;gap:8px;margin-top:4px;flex-wrap:wrap}
.ch-btn{flex:1;min-width:52px;padding:10px 4px;text-align:center;background:#0F1923;border:1.5px solid var(--border);border-radius:8px;cursor:pointer;font-size:.88rem;font-weight:700;color:var(--muted);transition:all .2s;user-select:none}
.ch-btn:hover{border-color:var(--accent);color:var(--accent)}
.ch-btn.selected{background:var(--accent);border-color:var(--accent);color:#fff}
.ch-sub{display:block;font-size:.6rem;margin-top:3px;font-weight:400;opacity:.8}
.ch-quick{display:flex;gap:6px;margin-top:7px;flex-wrap:wrap}
.ch-quick button{background:none;border:1px solid var(--border);border-radius:5px;color:var(--muted);font-size:.72rem;padding:4px 9px;cursor:pointer;transition:all .2s}
.ch-quick button:hover{border-color:var(--accent);color:var(--accent)}
.ch-summary{font-size:.78rem;color:var(--accent);margin-top:6px;min-height:1.2em}

/* Front matter toggles */
.fm-checks{display:flex;gap:10px;flex-wrap:wrap;margin-top:6px}
.fm-check{display:flex;align-items:center;gap:7px;background:#0F1923;border:1.5px solid var(--border);border-radius:8px;padding:9px 14px;cursor:pointer;user-select:none;transition:border .2s,background .2s}
.fm-check:hover{border-color:var(--accent)}
.fm-check.active{border-color:var(--accent);background:#152233}
.fm-check input[type=checkbox]{display:none}
.fm-check-icon{font-size:1rem}
.fm-check-label{font-size:.82rem;font-weight:600;color:var(--text)}
.fm-note{font-size:.73rem;color:var(--muted);margin-top:7px}

/* Custom TOC */
.toc-toggle{display:flex;align-items:center;gap:8px;cursor:pointer;margin-top:4px;user-select:none}
.toc-toggle input[type=checkbox]{width:16px;height:16px;accent-color:var(--accent);cursor:pointer}
.toc-toggle span{font-size:.82rem;color:var(--muted)}
#custom-toc-wrap{display:none;margin-top:10px}
#custom-toc{width:100%;background:#0F1923;border:1px solid var(--border);border-radius:8px;color:var(--text);font-size:.82rem;font-family:Consolas,monospace;padding:10px 12px;outline:none;resize:vertical;min-height:130px;transition:border .2s;line-height:1.6}
#custom-toc:focus{border-color:var(--accent)}
.toc-hint{font-size:.72rem;color:var(--muted);margin-top:5px;line-height:1.5}

/* Custom instructions */
.ci-toggle{display:flex;align-items:center;gap:8px;cursor:pointer;user-select:none}
.ci-toggle input[type=checkbox]{width:16px;height:16px;accent-color:var(--accent);cursor:pointer}
.ci-toggle span{font-size:.82rem;color:var(--muted)}
#ci-wrap{display:none;margin-top:10px}
#custom-instructions{width:100%;background:#0F1923;border:1px solid var(--border);border-radius:8px;color:var(--text);font-size:.85rem;padding:11px 14px;outline:none;resize:vertical;min-height:100px;transition:border .2s;line-height:1.65}
#custom-instructions:focus{border-color:var(--accent)}
#custom-instructions::placeholder{color:#3A5A7A}
.ci-hint{font-size:.72rem;color:var(--muted);margin-top:5px;line-height:1.5}

/* Level selector */
.level-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:4px}
.level-card{
  background:#0F1923;border:2px solid var(--border);border-radius:10px;
  padding:16px 14px;cursor:pointer;transition:border .2s,background .2s;text-align:center
}
.level-card:hover{border-color:var(--accent);background:#152233}
.level-card.selected{border-color:var(--accent);background:#152233}
.level-title{font-size:1rem;font-weight:700;color:var(--text);margin-bottom:6px}
.level-desc{font-size:.78rem;color:var(--muted);line-height:1.5}

/* Steps */
.steps{display:flex;counter-reset:step;margin-bottom:10px}
.step{flex:1;text-align:center;font-size:.78rem;color:var(--muted);padding:6px 4px;position:relative}
.step::before{counter-increment:step;content:counter(step);display:block;width:26px;height:26px;border-radius:50%;background:var(--border);color:var(--muted);font-weight:700;line-height:26px;margin:0 auto 5px;font-size:.8rem}
.step.active::before{background:var(--accent);color:#fff}
.step.done::before{background:var(--success);color:#fff;content:"✓"}

/* Buttons */
.btn{display:inline-flex;align-items:center;gap:8px;background:var(--accent);color:#fff;border:none;border-radius:9px;font-size:1rem;font-weight:600;padding:13px 28px;cursor:pointer;transition:background .2s;width:100%;justify-content:center}
.btn:hover{background:var(--accent2)}
.btn:disabled{background:#2A3F55;color:#4A6080;cursor:not-allowed}
.btn-success{background:var(--success)}
.btn-success:hover{background:#43A047}

/* Progress */
#progress-section{display:none}
.pbar-wrap{background:#0F1923;border-radius:999px;height:7px;margin:10px 0 18px;overflow:hidden}
.pbar{height:100%;background:var(--accent);border-radius:999px;animation:slide 1.4s ease-in-out infinite;width:35%}
@keyframes slide{0%{transform:translateX(-120%)}100%{transform:translateX(320%)}}

#log{background:#080F18;border:1px solid var(--border);border-radius:8px;font-family:Consolas,monospace;font-size:.82rem;line-height:1.6;color:#94B8D8;padding:14px 16px;height:280px;overflow-y:auto;white-space:pre-wrap}
.la{color:var(--accent)}.ls{color:var(--success);font-weight:600}.le{color:var(--error)}.lh{color:var(--accent);font-weight:700}

/* Download */
#download-section{display:none}
.dl-box{background:linear-gradient(135deg,#1A3025,#152B1E);border:1px solid #2E6040;border-radius:10px;padding:24px;text-align:center}
.dl-box .icon{font-size:2.6rem;margin-bottom:8px}
.dl-box p{color:#7AC99A;font-size:.88rem;margin-bottom:16px}

.email-note{background:#1A2A1A;border:1px solid #2E6040;border-radius:8px;padding:12px 16px;color:#7AC99A;font-size:.82rem;margin-top:12px;text-align:left}

@media(max-width:560px){.card{padding:20px 16px}.level-grid{grid-template-columns:1fr}.row-2{grid-template-columns:1fr}.chapter-grid{flex-wrap:wrap}.ch-btn{min-width:54px;flex:none}}
</style>
</head>
<body>

<header>
  <h1>📄 Academic Research Writeup Agent</h1>
  <p>Enter your details below — receive a complete Word document by email</p>
</header>

<main>

  <div class="steps" id="steps">
    <div class="step active" id="step1">Enter Details</div>
    <div class="step" id="step2">Generating</div>
    <div class="step" id="step3">Done</div>
  </div>

  <!-- ── Input form ── -->
  <div id="form-section">
    <div class="card">
      <h2>📝 Research Details</h2>

      <div class="form-group">
        <label>Project Topic</label>
        <input type="text" id="topic"
               placeholder="e.g. The Impact of AI on Healthcare Delivery in Sub-Saharan Africa"/>
      </div>

      <div class="form-group">
        <label>Research Level</label>
        <div class="level-grid">
          <div class="level-card selected" data-level="undergraduate" onclick="selectLevel(this)">
            <div class="level-title">🎓 Undergraduate</div>
            <div class="level-desc">Clear, well-structured analysis.<br>Accessible theory. Solid methodology.</div>
          </div>
          <div class="level-card" data-level="postgraduate" onclick="selectLevel(this)">
            <div class="level-title">🏛️ Postgraduate</div>
            <div class="level-desc">Critical engagement with theory.<br>Epistemological depth. Advanced methodology.</div>
          </div>
        </div>
      </div>

      <div class="form-group">
        <label>Chapters to Generate</label>
        <div class="chapter-grid">
          <div class="ch-btn selected" data-ch="1" onclick="toggleChapter(this)">1<span class="ch-sub">Intro</span></div>
          <div class="ch-btn selected" data-ch="2" onclick="toggleChapter(this)">2<span class="ch-sub">Lit. Rev.</span></div>
          <div class="ch-btn selected" data-ch="3" onclick="toggleChapter(this)">3<span class="ch-sub">Methodology</span></div>
          <div class="ch-btn selected" data-ch="4" onclick="toggleChapter(this)">4<span class="ch-sub">Results</span></div>
          <div class="ch-btn selected" data-ch="5" onclick="toggleChapter(this)">5<span class="ch-sub">Conclusions</span></div>
        </div>
        <div class="ch-quick">
          <button type="button" onclick="setChapters('all')">All</button>
          <button type="button" onclick="setChapters('1')">Ch 1 only</button>
          <button type="button" onclick="setChapters('1-3')">Ch 1–3</button>
          <button type="button" onclick="setChapters('3-5')">Ch 3–5</button>
          <button type="button" onclick="setChapters('none')">Clear</button>
        </div>
        <div class="ch-summary" id="ch-summary">Chapters: 1, 2, 3, 4, 5 (complete)</div>
      </div>

      <div class="form-group">
        <label>Front Matter Sections</label>
        <div class="fm-checks">
          <div class="fm-check active" data-fm="declaration" onclick="toggleFm(this)">
            <span class="fm-check-icon">📜</span>
            <span class="fm-check-label">Declaration</span>
          </div>
          <div class="fm-check active" data-fm="dedication" onclick="toggleFm(this)">
            <span class="fm-check-icon">❤️</span>
            <span class="fm-check-label">Dedication</span>
          </div>
          <div class="fm-check active" data-fm="acknowledgements" onclick="toggleFm(this)">
            <span class="fm-check-icon">🙏</span>
            <span class="fm-check-label">Acknowledgements</span>
          </div>
        </div>
        <div class="fm-note">ℹ️ Abstract is always included. Toggle any section above to include or exclude it.</div>
      </div>

      <div class="form-group">
        <label>Table of Contents</label>
        <label class="toc-toggle">
          <input type="checkbox" id="custom-toc-toggle" onchange="toggleCustomToc()"/>
          <span>Provide my own custom table of contents</span>
        </label>
        <div id="custom-toc-wrap">
          <textarea id="custom-toc" placeholder="Enter your TOC, one item per line. Example:&#10;CHAPTER ONE: INTRODUCTION&#10;  1.1  Background of the Study&#10;  1.2  Statement of the Problem&#10;  1.3  Research Objectives&#10;&#10;CHAPTER TWO: LITERATURE REVIEW&#10;  2.1  Conceptual Framework&#10;  ..."></textarea>
          <div class="toc-hint">💡 Each line becomes one TOC entry. Leave blank lines between chapters for spacing.</div>
        </div>
      </div>

      <div class="form-group">
        <label>Custom Instructions</label>
        <label class="ci-toggle">
          <input type="checkbox" id="ci-toggle-cb" onchange="toggleCi()"/>
          <span>Add custom instructions for document generation</span>
        </label>
        <div id="ci-wrap">
          <textarea id="custom-instructions"
            placeholder="Enter any specific instructions the AI should follow when writing your document.&#10;&#10;Examples:&#10;• Use Nigeria as the primary case study&#10;• Cite specific scholars or theoretical frameworks&#10;• Focus on quantitative methods only&#10;• Follow a particular institution's formatting guidelines"></textarea>
          <div class="ci-hint">💡 These instructions are applied to every chapter and front matter section generated.</div>
        </div>
      </div>

      <hr class="divider"/>

      <div class="row-2">
        <div class="form-group" style="margin-bottom:0">
          <label>Phone Number</label>
          <input type="tel" id="phone" placeholder="+234 800 000 0000"/>
        </div>
        <div class="form-group" style="margin-bottom:0">
          <label>Email Address</label>
          <input type="email" id="email" placeholder="client@example.com"/>
        </div>
      </div>
    </div>

    <div class="card" style="padding-top:20px;padding-bottom:20px">
      <button class="btn" id="gen-btn" onclick="start()">
        ⚡ Generate Research Document
      </button>
      <p style="color:var(--muted);font-size:.76rem;margin-top:10px;text-align:center">
        ~2–4 minutes per document · Finished document emailed on completion
      </p>
    </div>
  </div>

  <!-- ── Progress ── -->
  <div class="card" id="progress-section">
    <h2>⏳ Generating Your Document</h2>
    <div class="pbar-wrap"><div class="pbar" id="pbar"></div></div>
    <div id="log"></div>
  </div>

  <!-- ── Download ── -->
  <div id="download-section">
    <div class="dl-box">
      <div class="icon">✅</div>
      <h2 style="color:#4CAF50;margin-bottom:6px">Document Ready!</h2>
      <p id="dl-filename"></p>
      <a id="dl-link" href="#" download>
        <button class="btn btn-success" style="max-width:320px;margin:0 auto">
          📥 Download Word Document
        </button>
      </a>
      <div class="email-note" id="email-note">
        📧 The document has also been sent to the configured email address(es).
      </div>
    </div>
    <div style="text-align:center;margin-top:14px">
      <button onclick="reset()" style="background:var(--card);border:1px solid var(--border);color:var(--muted);border-radius:7px;padding:9px 20px;cursor:pointer;font-size:.88rem">
        ⬅ Generate Another
      </button>
    </div>
  </div>

</main>

<script>
let jobId=null, sse=null, selectedLevel='undergraduate';

function selectLevel(el){
  document.querySelectorAll('.level-card').forEach(c=>c.classList.remove('selected'));
  el.classList.add('selected');
  selectedLevel=el.dataset.level;
}

// ── Chapter toggle helpers ─────────────────────────────────
function getSelectedChapters(){
  return [...document.querySelectorAll('.ch-btn.selected')].map(b=>parseInt(b.dataset.ch)).sort((a,b)=>a-b);
}

function buildChaptersString(sel){
  // Compress [1,2,3,5] → "1-3,5"
  if(!sel.length) return '';
  const parts=[];let i=0;
  while(i<sel.length){
    let start=sel[i],end=sel[i];
    while(i+1<sel.length&&sel[i+1]===sel[i]+1){i++;end=sel[i];}
    parts.push(start===end?String(start):start+'-'+end);
    i++;
  }
  return parts.join(',');
}

function updateChSummary(){
  const sel=getSelectedChapters();
  const el=document.getElementById('ch-summary');
  if(!sel.length){el.textContent='No chapters selected';el.style.color='var(--error)';return;}
  // Build label
  const all=[1,2,3,4,5];
  let label='Ch ';
  if(JSON.stringify(sel)===JSON.stringify(all)){label='Chapters 1-5 (complete)';}
  else if(sel.length===1){label='Chapter '+sel[0]+' only';}
  else{
    // Detect contiguous range
    const isRange=sel.every((v,i)=>i===0||v===sel[i-1]+1);
    label='Chapters '+(isRange?sel[0]+'-'+sel[sel.length-1]:sel.join(', '));
  }
  el.textContent=label;
  el.style.color='var(--accent)';
}

function toggleChapter(el){
  el.classList.toggle('selected');
  updateChSummary();
}

function setChapters(preset){
  const btns=document.querySelectorAll('.ch-btn');
  btns.forEach(b=>{
    const n=parseInt(b.dataset.ch);
    if(preset==='all') b.classList.add('selected');
    else if(preset==='none') b.classList.remove('selected');
    else if(preset==='1') b.classList.toggle('selected',n===1);
    else if(preset==='1-3') b.classList.toggle('selected',n<=3);
    else if(preset==='3-5') b.classList.toggle('selected',n>=3);
  });
  updateChSummary();
}

function toggleFm(el){
  el.classList.toggle('active');
}

function getFmSections(){
  return [...document.querySelectorAll('.fm-check.active')].map(e=>e.dataset.fm);
}

function toggleCustomToc(){
  const show=document.getElementById('custom-toc-toggle').checked;
  document.getElementById('custom-toc-wrap').style.display=show?'block':'none';
}

function toggleCi(){
  const show=document.getElementById('ci-toggle-cb').checked;
  document.getElementById('ci-wrap').style.display=show?'block':'none';
}

function setStep(n){
  for(let i=1;i<=3;i++){
    const s=document.getElementById('step'+i);
    s.className='step'+(i<n?' done':i===n?' active':'');
  }
}

function log(msg,cls=''){
  const el=document.getElementById('log');
  el.innerHTML+=(cls?`<span class="${cls}">`+esc(msg)+'</span>':esc(msg))+'\n';
  el.scrollTop=el.scrollHeight;
}

function esc(s){return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}

async function start(){
  const topic=document.getElementById('topic').value.trim();
  const phone=document.getElementById('phone').value.trim();
  const email=document.getElementById('email').value.trim();
  if(!topic){alert('Please enter a research topic.');return;}

  document.getElementById('gen-btn').disabled=true;
  document.getElementById('gen-btn').textContent='⏳ Starting...';
  document.getElementById('form-section').style.display='none';
  document.getElementById('progress-section').style.display='block';
  document.getElementById('log').innerHTML='';
  setStep(2);

  const chArr=getSelectedChapters();
  if(!chArr.length){alert('Please select at least one chapter.');document.getElementById('form-section').style.display='block';document.getElementById('progress-section').style.display='none';document.getElementById('gen-btn').disabled=false;document.getElementById('gen-btn').textContent='⚡ Generate Research Document';setStep(1);return;}
  // Build compact string "1-5", "3", "1,3-5" — same format as PHP portal
  const chaptersStr=buildChaptersString(chArr);
  const customTocOn=document.getElementById('custom-toc-toggle').checked;
  const customToc=customTocOn?document.getElementById('custom-toc').value.trim():'';
  const fmSections=getFmSections();
  const payload={project_topic:topic,research_level:selectedLevel,chapters:chaptersStr,front_matter_sections:fmSections};
  if(customToc) payload.custom_toc=customToc;
  if(email) payload.email=email;
  if(phone) payload.phone=phone;
  const ciOn=document.getElementById('ci-toggle-cb').checked;
  const customInstructions=ciOn?document.getElementById('custom-instructions').value.trim():'';
  if(customInstructions) payload.custom_instructions=customInstructions;

  const res=await fetch('/generate',{
    method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify(payload)
  });
  const data=await res.json();
  if(data.error){log('Error: '+data.error,'le');return;}
  jobId=data.job_id;

  sse=new EventSource('/stream/'+jobId);
  let sseConnected=false;
  let pollTimeout;

  // Try SSE first
  sse.addEventListener('open',()=>sseConnected=true);

  sse.addEventListener('log',e=>{
    const d=JSON.parse(e.data);
    const cls={success:'ls',error:'le',accent:'la',header:'lh'}[d.tag]||'';
    log(d.msg,cls);
  });
  sse.addEventListener('done',e=>{
    sse.close();
    clearTimeout(pollTimeout);
    const d=JSON.parse(e.data);
    showDownload(d.filename,d.job_id,d.emailed);
  });
  sse.addEventListener('error_event',e=>{
    sse.close();
    clearTimeout(pollTimeout);
    log('\n❌ '+JSON.parse(e.data).msg,'le');
    document.getElementById('gen-btn').disabled=false;
    document.getElementById('gen-btn').textContent='⚡ Generate Research Document';
  });

  // Fallback: if SSE fails after 3 seconds, poll status instead
  sse.onerror=()=>{
    if(!sseConnected){
      sse.close();
      log('Connecting via polling...','la');
      pollJobStatus(jobId);
    }
  };

  setTimeout(()=>{
    if(!sseConnected){
      sse.close();
      log('Using polling for real-time updates...','la');
      pollJobStatus(jobId);
    }
  },3000);
}

function pollJobStatus(jobId){
  async function checkStatus(){
    try{
      const res=await fetch('/api/job-status/'+jobId);
      const job=await res.json();
      if(job.logs){
        job.logs.forEach(l=>{
          const cls={success:'ls',error:'le',accent:'la',header:'lh'}[l.tag]||'';
          log(l.msg,cls);
        });
      }
      if(job.status==='done'){
        showDownload(job.filename,jobId,job.emailed);
      }else if(job.status==='error'){
        log('\n❌ '+job.error,'le');
        document.getElementById('gen-btn').disabled=false;
        document.getElementById('gen-btn').textContent='⚡ Generate Research Document';
      }else{
        setTimeout(checkStatus,2000);
      }
    }catch(e){
      log('Status check failed: '+e.message,'le');
      setTimeout(checkStatus,5000);
    }
  }
  checkStatus();
}

function showDownload(filename,jid,emailed){
  document.getElementById('progress-section').style.display='none';
  document.getElementById('download-section').style.display='block';
  document.getElementById('dl-filename').textContent=filename;
  document.getElementById('dl-link').href='/download/'+jid;
  if(!emailed) document.getElementById('email-note').style.display='none';
  setStep(3);
}

function reset(){
  document.getElementById('download-section').style.display='none';
  document.getElementById('form-section').style.display='block';
  document.getElementById('gen-btn').disabled=false;
  document.getElementById('gen-btn').textContent='⚡ Generate Research Document';
  document.getElementById('topic').value='';
  document.getElementById('phone').value='';
  document.getElementById('email').value='';
  document.getElementById('custom-toc').value='';
  document.getElementById('custom-toc-toggle').checked=false;
  document.getElementById('custom-toc-wrap').style.display='none';
  document.getElementById('custom-instructions').value='';
  document.getElementById('ci-toggle-cb').checked=false;
  document.getElementById('ci-wrap').style.display='none';
  selectedLevel='undergraduate';
  document.querySelectorAll('.level-card').forEach((c,i)=>c.classList.toggle('selected',i===0));
  document.querySelectorAll('.fm-check').forEach(c=>c.classList.add('active'));
  setChapters('all');
  setStep(1);
}
</script>
</body>
</html>
"""


# ─────────────────────────────────────────────────────────
#  ROUTES
# ─────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/generate", methods=["POST"])
def generate():
    """
    POST /generate — JSON body fields  (all names match the web UI form)
    ─────────────────────────────────────────────────────────────────────
    Field             Type      Required  Description
    ─────────────────────────────────────────────────────────────────────
    project_topic     string    YES       The full research topic title
    research_level    string    YES       "undergraduate" or "postgraduate"
    chapters          string    no        Which chapters to write.
                                          Formats:  "3"        → chapter 3 only
                                                    "3-5"      → chapters 3,4,5
                                                    "1,3,5"    → chapters 1,3 and 5
                                                    "1,3-5"    → chapters 1,3,4,5
                                                    "1-5"/"all"→ all five chapters (default)
    custom_toc        string    no        Custom TOC text (one entry per line).
                                          Omit or leave blank to use the auto-generated TOC.
    email             string    no        Client email — document is sent here on completion.
    phone             string    no        Client phone number (stored for reference).
    ─────────────────────────────────────────────────────────────────────
    Returns: { "job_id": "uuid", "chapters": [1,2,3,4,5] }

    Example curl:
      curl -X POST http://localhost:5000/generate \\
           -H "Content-Type: application/json" \\
           -d '{"project_topic":"Impact of AI on Healthcare","research_level":"undergraduate","chapters":"3-5","email":"client@example.com"}'
    """
    data           = request.get_json(force=True, silent=True) or {}
    topic          = (data.get("project_topic") or data.get("topic") or "").strip()
    research_level = (data.get("research_level") or data.get("level") or "undergraduate").strip().lower()
    chapters_raw          = data.get("chapters")          # flexible — see parse_chapters()
    custom_toc            = (data.get("custom_toc") or "").strip() or None
    extra_email           = (data.get("email") or "").strip() or None
    front_matter_sections = data.get("front_matter_sections")  # list or None → defaults to all
    custom_instructions   = (data.get("custom_instructions") or "").strip() or None

    if not topic:
        return jsonify({"error": "project_topic is required"}), 400
    if research_level not in research_agent.LEVEL_PROFILES:
        research_level = "undergraduate"

    chapters_list = research_agent.parse_chapters(chapters_raw)

    job_id = str(uuid.uuid4())
    JOBS[job_id] = {
        "status":    "running",
        "log_queue": queue.Queue(),
        "file_path": None,
        "filename":  None,
        "emailed":   False,
        "error":     None,
    }

    threading.Thread(
        target=_run_agent,
        args=(job_id, topic, research_level, chapters_list,
              extra_email, custom_toc, front_matter_sections,
              custom_instructions),
        daemon=True
    ).start()

    return jsonify({"job_id": job_id, "chapters": chapters_list})


@app.route("/stream/<job_id>")
def stream(job_id):
    if job_id not in JOBS:
        abort(404)

    def generate_events():
        import json
        job = JOBS[job_id]
        while True:
            try:
                item = job["log_queue"].get(timeout=30)
            except queue.Empty:
                yield "event: ping\ndata: {}\n\n"
                continue

            if item["type"] == "log":
                yield f"event: log\ndata: {json.dumps({'msg': item['msg'], 'tag': item['tag']})}\n\n"
            elif item["type"] == "done":
                yield (f"event: done\ndata: {json.dumps({'filename': item['filename'], 'job_id': job_id, 'emailed': item['emailed']})}\n\n")
                break
            elif item["type"] == "error":
                yield f"event: error_event\ndata: {json.dumps({'msg': item['msg']})}\n\n"
                break

    return Response(generate_events(), mimetype="text/event-stream",
                    headers={
                        "Cache-Control": "no-cache, no-store, must-revalidate",
                        "Pragma": "no-cache",
                        "Expires": "0",
                        "X-Accel-Buffering": "no",
                        "Connection": "keep-alive",
                        "Access-Control-Allow-Origin": "*",
                        "Access-Control-Allow-Methods": "GET, OPTIONS",
                        "Access-Control-Allow-Headers": "Content-Type"
                    })


@app.route("/api/job-status/<job_id>")
def job_status(job_id):
    """Polling endpoint for job status (fallback if SSE fails)."""
    try:
        import json
        if job_id not in JOBS:
            return jsonify({"status": "error", "error": "Job not found"}), 404

        job = JOBS[job_id]
        logs = []

        # Drain all logs from queue without blocking
        try:
            while not job["log_queue"].empty():
                try:
                    item = job["log_queue"].get_nowait()
                    if item["type"] == "log":
                        logs.append({"msg": item["msg"], "tag": item["tag"]})
                    elif item["type"] == "done":
                        return jsonify({
                            "status": "done",
                            "filename": item["filename"],
                            "emailed": item["emailed"],
                            "logs": logs
                        })
                    elif item["type"] == "error":
                        return jsonify({
                            "status": "error",
                            "error": item["msg"],
                            "logs": logs
                        })
                except:
                    break
        except:
            pass

        return jsonify({
            "status": "generating",
            "logs": logs
        })
    except Exception as e:
        return jsonify({"status": "error", "error": str(e)}), 500


@app.route("/debug/jobs")
def debug_jobs():
    """Debug endpoint: show all jobs in memory"""
    jobs_info = {}
    for jid, job in JOBS.items():
        jobs_info[jid] = {
            "status": job.get("status"),
            "filename": job.get("filename"),
            "file_path": job.get("file_path"),
            "exists": os.path.exists(job.get("file_path", "")) if job.get("file_path") else False
        }
    return jsonify(jobs_info)


@app.route("/download/<job_id>")
def download(job_id):
    if job_id not in JOBS:
        abort(404)
    path = JOBS[job_id].get("file_path")
    if not path or not os.path.exists(path):
        abort(404)
    return send_file(
        path,
        as_attachment=True,
        download_name=JOBS[job_id]["filename"],
        mimetype=(
            "application/vnd.openxmlformats-officedocument"
            ".wordprocessingml.document"
        )
    )


# ─────────────────────────────────────────────────────────
#  BACKGROUND AGENT RUNNER
# ─────────────────────────────────────────────────────────

def _run_agent(job_id: str, topic: str, research_level: str,
               chapters_list: list = None,
               extra_email: str = None,
               custom_toc: str = None,
               front_matter_sections: list = None,
               custom_instructions: str = None):
    job = JOBS[job_id]
    q   = job["log_queue"]

    chapters_list = chapters_list or list(range(1, 6))

    def log(msg, tag="info"):
        q.put({"type": "log", "msg": msg, "tag": tag})

    try:
        os.environ["ANTHROPIC_API_KEY"] = config.ANTHROPIC_API_KEY
        client = research_agent.anthropic.Anthropic(
            api_key=config.ANTHROPIC_API_KEY
        )
        level_label  = research_agent.LEVEL_PROFILES[research_level]["label"]
        ch_label     = research_agent.fmt_chapters_label(chapters_list)
        toc_src      = "custom" if custom_toc else "auto-generated"

        # Resolve which optional front matter sections to include
        _fm_all = ["declaration", "dedication", "acknowledgements"]
        if front_matter_sections is None:
            fm_include = _fm_all
        else:
            fm_include = [s.lower().strip() for s in front_matter_sections
                          if s.lower().strip() in _fm_all]
        fm_label = ", ".join(s.title() for s in fm_include) + ", Abstract" if fm_include else "Abstract only"

        log("=" * 52, "header")
        log(f"  TOPIC    : {topic}", "header")
        log(f"  LEVEL    : {level_label}", "header")
        log(f"  CHAPTERS : {ch_label}", "header")
        log(f"  FRONT    : {fm_label}", "header")
        log(f"  TOC      : {toc_src}", "header")
        log(f"  MODEL    : {config.MODEL}", "header")
        if custom_instructions:
            ci_preview = custom_instructions[:60] + ("…" if len(custom_instructions) > 60 else "")
            log(f"  CUSTOM   : {ci_preview}", "header")
        log("=" * 52, "header")
        log("")

        # Front matter
        fm_parts = fm_label
        log(f"► Front matter ({fm_label})...", "accent")
        front = research_agent.generate_front_matter(
            client, topic, research_level, model=config.MODEL,
            front_matter_sections=fm_include,   # pass [] when all deselected, not None (None = include all)
            custom_instructions=custom_instructions
        )
        log(f"  ✓ Front matter — {len(front):,} chars", "success")
        log("")

        # Generate requested chapters
        # [TEMPORARY] Skipping chapters to save tokens during download troubleshooting
        chapters = {}
        # for num in chapters_list:
        #     name = research_agent.CHAPTER_SUBTITLES[num]
        #     log(f"► Chapter {num}: {name}...", "accent")
        #     chapters[num] = research_agent.generate_chapter(
        #         client, topic, num, research_level, model=config.MODEL,
        #         custom_instructions=custom_instructions
        #     )
        #     log(f"  ✓ Chapter {num} complete — {len(chapters[num]):,} chars", "success")
        #     log("")

        # ── Extract ## REFERENCES from chapter 5 (Issue 3) ──────────────
        references_text = ""
        if 5 in chapters:
            ref_match = re.search(r'\n## REFERENCES', chapters[5], re.IGNORECASE)
            if ref_match:
                references_text = chapters[5][ref_match.start():]
                chapters[5]     = chapters[5][:ref_match.start()]

        # Build document
        log("► Assembling Word document...", "accent")
        safe     = re.sub(r"[^\w\s-]", "", topic).strip().replace(" ", "_")[:50]
        filename = f"Research_{safe}.docx"
        out_path = os.path.join(OUTPUT_DIR, f"{job_id}_{filename}")
        log(f"  Saving to: {out_path}", "info")

        doc = research_agent.Document()
        research_agent.set_document_defaults(doc)
        research_agent.build_title_page(doc, topic, research_level)
        research_agent.build_front_matter_page(doc, front)
        research_agent.build_toc_page(doc, research_level,
                                      chapters_list=chapters_list,
                                      custom_toc=custom_toc,
                                      front_matter_sections=fm_include)  # pass [] not None
        research_agent.build_abbreviations_page(doc)

        # One FootnoteManager per document — shared across all chapters
        fn_mgr = research_agent.FootnoteManager(doc)

        # [TEMPORARY] Skipping chapter pages to save tokens during download troubleshooting
        # for num in chapters_list:
        #     research_agent.build_chapter_page(doc, num, chapters[num], fn_mgr=fn_mgr)

        # Dedicated references page (always after chapters, never inside them)
        if references_text.strip():
            research_agent.build_references_page(doc, references_text)

        doc.save(out_path)

        # Inject footnotes.xml into the zip AFTER saving (zip-injection approach
        # avoids python-docx internal API issues with XmlPart._element)
        fn_mgr.inject(out_path)

        job["file_path"] = out_path
        job["filename"]  = filename
        job["status"]    = "done"

        log(f"  ✓ Document saved: {filename}", "success")
        log("")

        # Send email (to config recipients + optional per-job client email)
        log("► Sending email...", "accent")
        send_email(
            file_path=out_path,
            filename=filename,
            topic=topic,
            research_level=research_level,
            log_fn=lambda m: log(m, "info"),
            extra_recipients=[extra_email] if extra_email else None,
        )
        job["emailed"] = True
        log("")
        log("=" * 52, "header")
        log("  ✅  ALL DONE — document ready for download!", "success")
        log("=" * 52, "header")

        q.put({"type": "done", "filename": filename, "emailed": job["emailed"]})

    except Exception as exc:
        import traceback
        job["status"] = "error"
        job["error"]  = str(exc)
        log(f"\n❌ Error: {exc}", "error")
        log(traceback.format_exc(), "error")
        q.put({"type": "error", "msg": str(exc)})


# ─────────────────────────────────────────────────────────
#  LAUNCH
# ─────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print("=" * 56)
    print("  Academic Research Agent — Web Interface")
    print(f"  http://localhost:{port}")
    print("=" * 56)
    app.run(debug=False, host="0.0.0.0", port=port, threaded=True)

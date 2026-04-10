/* ================================================================
   CheckForm KLTN — Frontend JavaScript
   ================================================================ */

// ── State ────────────────────────────────────────────────────────
let pendingFiles = [];   // File objects
let allResults   = [];   // {filename, score, ...}
let selectedIdx  = -1;
let currentTab   = 'ERROR';

// CAPTCHA state
let captchaToken  = '';   // Token từ server
let captchaValid  = false; // Đã xác nhận đúng chưa

// ── DOM Refs ─────────────────────────────────────────────────────
const dropZone      = document.getElementById('drop-zone');
const fileInput     = document.getElementById('file-input');
const fileQueue     = document.getElementById('file-queue');
const queueTitle    = document.getElementById('queue-title');
const queueList     = document.getElementById('queue-list');
const btnSelectFiles= document.getElementById('btn-select-files');
const btnClearQueue = document.getElementById('btn-clear-queue');
const btnRun        = document.getElementById('btn-run');
const progressSec   = document.getElementById('progress-section');
const resultsSec    = document.getElementById('results-section');
const summaryRow    = document.getElementById('summary-row');
const resultsTbody  = document.getElementById('results-tbody');
const detailPanel   = document.getElementById('detail-panel');
const progressBar   = document.getElementById('progress-bar');
const progressLabel = document.getElementById('progress-label');
const progressCounter= document.getElementById('progress-counter');
const progressFile  = document.getElementById('progress-file');
const configModal   = document.getElementById('config-modal');
const advisorsTa    = document.getElementById('advisors-textarea');
const toastWrap     = document.getElementById('toast-wrap');
const captchaBox    = document.getElementById('captcha-box');
const captchaQEl    = document.getElementById('captcha-question');
const captchaInput  = document.getElementById('captcha-input');
const captchaHint   = document.getElementById('captcha-hint');

// ── CAPTCHA ───────────────────────────────────────────────────
async function fetchCaptcha() {
  captchaToken = '';
  captchaValid = false;
  btnRun.disabled = true;
  captchaInput.value = '';
  captchaInput.className = 'captcha-input';
  captchaHint.textContent = '';
  captchaHint.className = 'captcha-hint';
  captchaQEl.textContent = 'Đang tải...';
  try {
    const r = await fetch('/api/captcha');
    if (!r.ok) throw new Error('Server error');
    const d = await r.json();
    captchaToken = d.token;
    captchaQEl.textContent = d.question;
  } catch {
    captchaQEl.textContent = 'Lỗi tải CAPTCHA — nhấn ↻ để thử lại';
  }
}

function checkCaptchaInput() {
  const val = captchaInput.value.trim();
  if (!val) {
    captchaInput.className = 'captcha-input';
    captchaHint.textContent = '';
    captchaHint.className = 'captcha-hint';
    captchaValid = false;
    btnRun.disabled = true;
    return;
  }
  // Kiểm tra real-time với server
  fetch('/api/captcha/check', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ token: captchaToken, answer: val }),
  }).then(r => r.json()).then(d => {
    if (d.valid) {
      captchaInput.className = 'captcha-input correct';
      captchaHint.textContent = '✅ Xác nhận đúng!';
      captchaHint.className = 'captcha-hint ok';
      captchaValid = true;
      btnRun.disabled = false;
    } else {
      captchaInput.className = 'captcha-input wrong';
      captchaHint.textContent = '❌ Không đúng, thử lại.';
      captchaHint.className = 'captcha-hint err';
      captchaValid = false;
      btnRun.disabled = true;
    }
  }).catch(() => {});
}

captchaInput.addEventListener('input', checkCaptchaInput);
captchaInput.addEventListener('keydown', e => { if (e.key === 'Enter' && captchaValid) runCheck(); });
document.getElementById('btn-captcha-refresh').addEventListener('click', fetchCaptcha);

// ── UPLOAD / DRAG & DROP ─────────────────────────────────────────
btnSelectFiles.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('click', e => {
  if (e.target === dropZone || e.target.closest('.upload-icon') || e.target.tagName === 'H2' || e.target.tagName === 'P') {
    fileInput.click();
  }
});

dropZone.addEventListener('dragover', e => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});
dropZone.addEventListener('dragleave', e => {
  if (!dropZone.contains(e.relatedTarget)) dropZone.classList.remove('drag-over');
});
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const files = [...e.dataTransfer.files].filter(f => f.name.endsWith('.docx'));
  if (!files.length) { toast('Chỉ chấp nhận file .docx', 'error'); return; }
  addFilesToQueue(files);
});

fileInput.addEventListener('change', () => {
  addFilesToQueue([...fileInput.files]);
  fileInput.value = '';
});

function addFilesToQueue(files) {
  const existing = new Set(pendingFiles.map(f => `${f.name}_${f.size}`));
  const added = [];
  for (const f of files) {
    if (!f.name.endsWith('.docx')) continue;
    const key = `${f.name}_${f.size}`;
    if (!existing.has(key)) { pendingFiles.push(f); existing.add(key); added.push(f); }
  }
  renderQueue();
}

function renderQueue() {
  queueList.innerHTML = '';
  pendingFiles.forEach((f, i) => {
    const item = document.createElement('div');
    item.className = 'queue-item';
    item.innerHTML = `
      <div class="queue-file-icon">📄</div>
      <div class="queue-name">${esc(f.name)}</div>
      <div class="queue-size">${fmtSize(f.size)}</div>
      <button class="queue-remove" data-idx="${i}" title="Xóa">×</button>
    `;
    queueList.appendChild(item);
  });
  queueList.querySelectorAll('.queue-remove').forEach(btn =>
    btn.addEventListener('click', () => { pendingFiles.splice(+btn.dataset.idx, 1); renderQueue(); })
  );
  queueTitle.textContent = `${pendingFiles.length} file đã chọn`;
  if (pendingFiles.length > 0) {
    fileQueue.style.display = 'block';
    document.getElementById('upload-section').querySelector('.upload-zone').style.display = 'none';
    // Tải CAPTCHA mới mỗi khi queue thay đổi từ 0 → có file
    if (!captchaToken) fetchCaptcha();
  } else {
    fileQueue.style.display = 'none';
    captchaToken = ''; captchaValid = false; btnRun.disabled = true;
    document.getElementById('upload-section').querySelector('.upload-zone').style.display = '';
  }
}

btnClearQueue.addEventListener('click', () => { pendingFiles = []; renderQueue(); });

// ── RUN CHECK ────────────────────────────────────────────────────
btnRun.addEventListener('click', runCheck);

async function runCheck() {
  if (!pendingFiles.length) { toast('Chưa có file nào', 'error'); return; }
  if (!captchaValid) { toast('Vui lòng giải câu đố xác nhận trước', 'error'); return; }

  // Hiện progress
  allResults = [];
  resultsSec.style.display = 'none';
  progressSec.style.display = 'block';
  progressBar.style.width = '0%';
  progressLabel.textContent = 'Đang gửi file...';
  progressCounter.textContent = '';
  progressFile.textContent = '';

  const fd = new FormData();
  pendingFiles.forEach(f => fd.append('files', f));
  // Đính kèm captcha token và đáp án
  fd.append('captcha_token',  captchaToken);
  fd.append('captcha_answer', captchaInput.value.trim());
  // Reset captcha sau khi gửi (token 1 lần)
  captchaToken = ''; captchaValid = false; btnRun.disabled = true;

  let resp;
  try {
    resp = await fetch('/api/check', { method: 'POST', body: fd });
  } catch (e) {
    toast('Lỗi kết nối server: ' + e.message, 'error');
    progressSec.style.display = 'none';
    return;
  }
  const data = await resp.json();
  if (data.error) { toast(data.error, 'error'); progressSec.style.display = 'none'; return; }

  if (!data.streaming) {
    // 1 file — kết quả ngay
    allResults = data.results || [];
    progressSec.style.display = 'none';
    showResults();
  } else {
    // Nhiều file — SSE
    const jobId = data.job_id;
    const total = data.total;
    progressLabel.textContent = `Đang kiểm tra ${total} file...`;

    const evtSource = new EventSource(`/api/stream/${jobId}`);
    evtSource.onmessage = e => {
      const msg = JSON.parse(e.data);
      if (msg.done) {
        evtSource.close();
        progressSec.style.display = 'none';
        showResults();
      } else if (msg.result) {
        allResults.push(msg.result);
        const pct = Math.round(msg.done_count / total * 100);
        progressBar.style.width = pct + '%';
        progressCounter.textContent = `${msg.done_count}/${total}`;
        progressFile.textContent = msg.result.filename;
        // Stream: hiện bảng sớm
        if (allResults.length === 1) {
          resultsSec.style.display = 'block';
          renderSummary();
          renderTable();
        } else {
          renderSummary();
          appendTableRow(allResults.length - 1, msg.result);
        }
      }
    };
    evtSource.onerror = () => {
      evtSource.close();
      progressSec.style.display = 'none';
      if (allResults.length) showResults();
      else toast('Lỗi kết nối stream', 'error');
    };
  }
}

// ── RESULTS ──────────────────────────────────────────────────────
function showResults() {
  resultsSec.style.display = 'block';
  renderSummary();
  renderTable();
  pendingFiles = [];
  renderQueue();
}

function renderSummary() {
  const total   = allResults.length;
  const passed  = allResults.filter(r => r.score >= 70).length;
  const errors  = allResults.reduce((s, r) => s + (r.error_count || 0), 0);
  const avgScore= total ? Math.round(allResults.reduce((s, r) => s + r.score, 0) / total) : 0;

  summaryRow.innerHTML = `
    ${card('📊', total,          'Tổng file',     '#8b949e', 'rgba(139,148,158,0.1)')}
    ${card('✅', passed,         'Đạt (≥70đ)',   '#10b981', 'rgba(16,185,129,0.1)')}
    ${card('❌', total - passed, 'Không đạt',    '#ef4444', 'rgba(239,68,68,0.1)')}
    ${card('⚠️', errors,         'Tổng lỗi',     '#f59e0b', 'rgba(245,158,11,0.1)')}
    ${card('🎯', avgScore + '/100','Điểm TB',    '#2563eb', 'rgba(37,99,235,0.1)')}
  `;
}

function card(icon, val, lbl, color, bg) {
  return `<div class="summary-card">
    <div class="summary-icon" style="background:${bg}">${icon}</div>
    <div>
      <div class="summary-val" style="color:${color}">${val}</div>
      <div class="summary-lbl">${lbl}</div>
    </div>
  </div>`;
}

function renderTable() {
  resultsTbody.innerHTML = '';
  if (!allResults.length) {
    resultsTbody.innerHTML = `<tr class="empty-row"><td colspan="8">Không có kết quả</td></tr>`;
    return;
  }
  allResults.forEach((r, i) => appendTableRow(i, r));
}

function appendTableRow(i, r) {
  // Remove empty row if any
  const empty = resultsTbody.querySelector('.empty-row');
  if (empty) empty.remove();

  const tr = document.createElement('tr');
  tr.dataset.idx = i;
  if (i === selectedIdx) tr.classList.add('selected');
  tr.innerHTML = `
    <td class="td-name"><span title="${esc(r.filename)}">${esc(r.filename)}</span></td>
    <td class="td-muted">${esc(r.student_name || '—')}</td>
    <td class="td-center td-muted">${esc(r.student_id || '—')}</td>
    <td class="td-muted" style="font-size:12px">${esc((r.advisor || '—').substring(0,20))}</td>
    <td class="td-center"><span style="color:${r.error_count ? 'var(--red)' : 'var(--text3)'};font-weight:600">${r.error_count || 0}</span></td>
    <td class="td-center"><span style="color:${r.warn_count ? 'var(--yellow)' : 'var(--text3)'};font-weight:600">${r.warn_count || 0}</span></td>
    <td class="td-center"><span class="score-badge" style="color:${scoreColor(r.score)}">${r.score}</span></td>
    <td class="td-center">${gradePill(r.letter_grade)}</td>
  `;
  tr.addEventListener('click', () => selectResult(i));
  resultsTbody.appendChild(tr);
}

function selectResult(i) {
  selectedIdx = i;
  resultsTbody.querySelectorAll('tr').forEach((tr, j) => {
    tr.classList.toggle('selected', +tr.dataset.idx === i);
  });
  showDetail(allResults[i]);
}

function scoreColor(s) {
  if (s >= 70) return 'var(--green)';
  if (s >= 50) return 'var(--yellow)';
  return 'var(--red)';
}

function gradePill(g) {
  if (!g) return '—';
  const cls = g.startsWith('A') ? 'grade-A' : g.startsWith('B') ? 'grade-B' : g.startsWith('C') ? 'grade-C' : 'grade-DE';
  return `<span class="grade-pill ${cls}">${esc(g)}</span>`;
}

// ── DETAIL PANEL ─────────────────────────────────────────────────
function showDetail(r) {
  const score  = r.score;
  const color  = scoreColor(score);
  const issues = r.issues || [];
  const errors = issues.filter(i => i.severity === 'ERROR');
  const warns  = issues.filter(i => i.severity === 'WARNING');
  const infos  = issues.filter(i => i.severity === 'INFO');

  // Verdict
  let verdict, verdictColor;
  if (score >= 90)       { verdict = '✅ Đạt tốt';    verdictColor = 'rgba(16,185,129,0.15)'; }
  else if (score >= 70)  { verdict = '✔ Đạt';          verdictColor = 'rgba(16,185,129,0.1)'; }
  else if (score >= 50)  { verdict = '⚠ Cần sửa';     verdictColor = 'rgba(245,158,11,0.1)'; }
  else                   { verdict = '❌ Không đạt';   verdictColor = 'rgba(239,68,68,0.1)'; }

  // SVG ring
  const circ = 2 * Math.PI * 26;
  const filled = circ * (1 - score / 100);
  const ringHtml = `
    <div class="score-ring">
      <svg viewBox="0 0 64 64">
        <circle cx="32" cy="32" r="26" fill="none" stroke="rgba(255,255,255,0.06)" stroke-width="6"/>
        <circle cx="32" cy="32" r="26" fill="none" stroke="${color}" stroke-width="6"
          stroke-dasharray="${circ}" stroke-dashoffset="${filled}"
          stroke-linecap="round" style="transition:stroke-dashoffset 0.6s ease"/>
      </svg>
      <div class="score-ring-text" style="color:${color}">
        <span>${score}</span>
        <span class="score-ring-sub">/100</span>
      </div>
    </div>`;

  detailPanel.innerHTML = `
    <div class="detail-header">
      <div class="detail-score-row">
        ${ringHtml}
        <div class="detail-info">
          <div class="detail-name">${esc(r.filename)}</div>
          <div class="detail-meta">
            <span>Sinh viên: </span>${esc(r.student_name || '—')}<br>
            <span>MSSV: </span>${esc(r.student_id || '—')}<br>
            <span>GVHD: </span>${esc((r.advisor || '—').substring(0, 40))}<br>
            <span>Đề tài: </span><span title="${esc(r.title || '')}">${esc((r.title || '—').substring(0, 50))}${r.title && r.title.length > 50 ? '…' : ''}</span>
          </div>
          <div class="detail-verdict" style="color:${color};background:${verdictColor}">${verdict} · ${esc(r.letter_grade || '')}</div>
        </div>
      </div>
    </div>
    <div class="detail-tabs" id="detail-tabs">
      <button class="detail-tab ${currentTab==='ERROR'?'active':''}" data-tab="ERROR">
        Lỗi <span class="tab-count tab-count-error">${errors.length}</span>
      </button>
      <button class="detail-tab ${currentTab==='WARNING'?'active':''}" data-tab="WARNING">
        Cảnh báo <span class="tab-count tab-count-warning">${warns.length}</span>
      </button>
      <button class="detail-tab ${currentTab==='INFO'?'active':''}" data-tab="INFO">
        Thông tin <span class="tab-count tab-count-info">${infos.length}</span>
      </button>
    </div>
    <div class="detail-issues" id="detail-issues"></div>
  `;

  // Tab switching
  detailPanel.querySelectorAll('.detail-tab').forEach(btn => {
    btn.addEventListener('click', () => {
      currentTab = btn.dataset.tab;
      detailPanel.querySelectorAll('.detail-tab').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      renderIssueList(issues);
    });
  });

  renderIssueList(issues);
}

function renderIssueList(issues) {
  const container = document.getElementById('detail-issues');
  if (!container) return;
  const filtered = issues.filter(i => i.severity === currentTab);
  container.innerHTML = '';
  if (!filtered.length) {
    container.innerHTML = `<div class="detail-no-issues">✓ Không có mục nào trong nhóm này.</div>`;
    return;
  }
  filtered.forEach(iss => {
    const card = document.createElement('div');
    card.className = `issue-card issue-card-${iss.severity}`;
    const icon = iss.severity === 'ERROR' ? '❌' : iss.severity === 'WARNING' ? '⚠️' : 'ℹ️';
    card.innerHTML = `
      <div class="issue-cat issue-cat-${iss.severity}">
        ${icon} ${esc(iss.category)}
        ${iss.location ? `<span class="issue-loc">[${esc(iss.location.substring(0,40))}]</span>` : ''}
      </div>
      <div class="issue-msg">${esc(iss.message)}</div>
      ${iss.suggestion ? `<div class="issue-sug">→ ${esc(iss.suggestion)}</div>` : ''}
    `;
    container.appendChild(card);
  });
}

// ── NEW CHECK ─────────────────────────────────────────────────────
document.getElementById('btn-new-check').addEventListener('click', () => {
  allResults = [];
  selectedIdx = -1;
  resultsSec.style.display = 'none';
  detailPanel.innerHTML = `<div class="detail-empty">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2"/></svg>
    <p>Chọn một file để xem chi tiết</p>
  </div>`;
  dropZone.style.display = '';
});

// ── EXPORT EXCEL ──────────────────────────────────────────────────
document.getElementById('btn-export').addEventListener('click', () => {
  if (!allResults.length) { toast('Chưa có dữ liệu', 'error'); return; }

  // Dùng hidden form POST thay vì fetch + blob URL
  // → Safari/Chrome/Firefox đều nhận đúng tên file từ Content-Disposition
  const form = document.createElement('form');
  form.method  = 'POST';
  form.action  = '/api/export';
  form.style.display = 'none';

  const input = document.createElement('input');
  input.type  = 'hidden';
  input.name  = 'payload';
  input.value = JSON.stringify({ results: allResults });

  form.appendChild(input);
  document.body.appendChild(form);
  form.submit();
  document.body.removeChild(form);

  toast('✅ Đang tải xuống Excel...', 'success');
});

// ── CONFIG MODAL ─────────────────────────────────────────────────
document.getElementById('btn-config').addEventListener('click', async () => {
  const cfg = await fetch('/api/config').then(r => r.json()).catch(() => ({ advisors: [] }));
  advisorsTa.value = (cfg.advisors || []).join('\n');
  configModal.style.display = 'flex';
});
document.getElementById('btn-cancel-config').addEventListener('click', () => {
  configModal.style.display = 'none';
});
document.getElementById('modal-close').addEventListener('click', () => {
  configModal.style.display = 'none';
});
configModal.addEventListener('click', e => {
  if (e.target === configModal) configModal.style.display = 'none';
});

document.getElementById('btn-save-config').addEventListener('click', async () => {
  const lines = advisorsTa.value.split('\n').map(s => s.trim()).filter(Boolean);
  try {
    const r = await fetch('/api/config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ advisors: lines }),
    });
    const d = await r.json();
    if (d.ok) {
      toast(`✅ Đã lưu ${d.count} giảng viên`, 'success');
      configModal.style.display = 'none';
    } else {
      toast('Lỗi lưu cấu hình', 'error');
    }
  } catch(e) {
    toast('Lỗi: ' + e.message, 'error');
  }
});

// ── TOAST ─────────────────────────────────────────────────────────
function toast(msg, type = 'info') {
  const el = document.createElement('div');
  el.className = `toast toast-${type}`;
  el.textContent = msg;
  toastWrap.appendChild(el);
  setTimeout(() => { el.style.transition = '0.3s'; el.style.opacity = '0'; setTimeout(() => el.remove(), 300); }, 3500);
}

// ── HELPERS ───────────────────────────────────────────────────────
function esc(s) {
  if (s == null) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function fmtSize(b) {
  if (b < 1024) return b + ' B';
  if (b < 1024*1024) return (b/1024).toFixed(1) + ' KB';
  return (b/1024/1024).toFixed(1) + ' MB';
}

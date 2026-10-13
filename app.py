<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Ticket Reconciliation Portal</title>
<style>
  body { margin:0; min-height:100vh; background:#070b12; color:#e6edf3;
         font-family: "Segoe UI", Arial, sans-serif; display:flex; justify-content:center; padding:40px 16px; box-sizing:border-box; }
  .card { width:100%; max-width:560px; background:#0d131c; border:1px solid #1f2a37; border-radius:14px; padding:28px; }
  h1 { font-size:22px; margin:0 0 6px; }
  .sub { color:#8b98a9; font-size:14px; margin-bottom:24px; }
  label.section { display:block; font-size:12px; font-weight:700; letter-spacing:.04em; margin:18px 0 8px; }
  .mis { color:#5b9bff; } .pdf { color:#3fcf8e; }
  .picker { border:1px solid #243041; border-radius:10px; padding:12px; display:flex; align-items:center; gap:12px; background:#0a0f17; }
  .picker button { background:#1d4ed8; color:#fff; border:0; border-radius:8px; padding:9px 14px; cursor:pointer; font-weight:600; }
  .picker.pdf button { background:#15803d; }
  .picker span { color:#8b98a9; font-size:14px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  input[type=file] { display:none; }
  #run { width:100%; margin-top:24px; padding:14px; border:0; border-radius:10px; background:#1d4ed8; color:#fff;
         font-size:16px; font-weight:700; cursor:pointer; }
  #run:disabled { background:#1b2b52; color:#7a88a0; cursor:not-allowed; }
  #status { margin-top:18px; font-size:14px; line-height:1.5; color:#b8c4d4; min-height:22px; white-space:pre-wrap; }
  #status.error { color:#ff8a8a; }
  #status.done { color:#3fcf8e; }
</style>
</head>
<body>
<div class="card">
  <h1>Ticket Reconciliation Portal</h1>
  <div class="sub">Upload the MIS report and all ticket PDFs (any number) to get one audited sheet.</div>

  <label class="section mis">1. MIS REPORT (.XLSX)</label>
  <div class="picker">
    <button type="button" onclick="document.getElementById('mis').click()">Choose File</button>
    <span id="misName">No file chosen</span>
    <input type="file" id="mis" accept=".xlsx">
  </div>

  <label class="section pdf">2. PDF TICKETS (.PDF)</label>
  <div class="picker pdf">
    <button type="button" onclick="document.getElementById('pdfs').click()">Choose Files</button>
    <span id="pdfName">No files chosen</span>
    <input type="file" id="pdfs" accept=".pdf" multiple>
  </div>

  <button id="run" disabled>Run audit</button>
  <div id="status"></div>
</div>

<script>
const misInput = document.getElementById('mis');
const pdfInput = document.getElementById('pdfs');
const runBtn = document.getElementById('run');
const statusBox = document.getElementById('status');

function refresh() {
  document.getElementById('misName').textContent = misInput.files[0] ? misInput.files[0].name : 'No file chosen';
  const n = pdfInput.files.length;
  document.getElementById('pdfName').textContent = n ? `${n} file${n === 1 ? '' : 's'} selected` : 'No files chosen';
  runBtn.disabled = !(misInput.files.length && n);
}
misInput.addEventListener('change', refresh);
pdfInput.addEventListener('change', refresh);

function show(text, cls) {
  statusBox.textContent = text;
  statusBox.className = cls || '';
}

async function runAudit() {
  runBtn.disabled = true;
  show('Uploading files...');

  const form = new FormData();
  form.append('mis_file', misInput.files[0]);
  for (const f of pdfInput.files) form.append('pdf_files', f);

  let job;
  try {
    const res = await fetch('/api/audit', { method: 'POST', body: form });
    job = await res.json();
    if (!res.ok) throw new Error(job.details || job.error || `Server returned ${res.status}`);
  } catch (err) {
    show('Upload failed: ' + err.message, 'error');
    runBtn.disabled = false;
    return;
  }

  const jobId = job.job_id;
  const poll = async () => {
    try {
      const res = await fetch('/api/status/' + jobId);
      const s = await res.json();
      if (s.state === 'running') {
        show(`Auditing ${s.pdf_count} tickets... ${s.elapsed_seconds}s elapsed.\nYou can keep this page open.`);
        setTimeout(poll, 3000);
      } else if (s.state === 'done') {
        show('Audit finished. Downloading the sheet...', 'done');
        window.location = '/api/download/' + jobId;
        runBtn.disabled = false;
      } else {
        show('Audit failed: ' + (s.error || 'unknown error'), 'error');
        runBtn.disabled = false;
      }
    } catch (err) {
      show('Lost contact with the server, retrying...', 'error');
      setTimeout(poll, 5000);
    }
  };
  poll();
}

runBtn.addEventListener('click', runAudit);
</script>
</body>
</html>

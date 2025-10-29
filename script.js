/*
  Features:
  - sheets: {name, priority, comments, data[][]}
  - autosave localStorage
  - save to GitHub via REST API (requires token input)
  - add hover + glow effects via CSS
*/

/* Storage and default */
const STORAGE_KEY = 'localExcel.github.v1';
let state = { sheets: [], active: 0 };
function defaultSheet(name='Sheet1', rows=10, cols=6){
  const data = []; for(let r=0;r<rows;r++) data.push(new Array(cols).fill(''));
  return { name, priority:'Normal', comments:'', data };
}

function loadState(){
  try{
    const raw = localStorage.getItem(STORAGE_KEY);
    if(raw){ state = JSON.parse(raw); }
    if(!state.sheets || state.sheets.length===0) state = { sheets:[defaultSheet('Sheet1')], active:0 };
  }catch(e){ console.error(e); state = { sheets:[defaultSheet('Sheet1')], active:0 }; }
}
function saveState(){
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  document.getElementById('saveStatus').textContent = 'Saved (local) ' + new Date().toLocaleTimeString();
}

/* Render UI */
const sheetsEl = document.getElementById('sheets');
const tableEl = document.getElementById('sheetTable');
function renderSheetsList(){
  sheetsEl.innerHTML = '';
  state.sheets.forEach((s,i)=>{
    const btn = document.createElement('div');
    btn.className = 'sheet-btn' + (i===state.active ? ' active':'');
    btn.innerHTML = `<div style="font-weight:600">${s.name}</div><div class="muted">${s.priority}</div>`;
    btn.onclick = ()=>{ state.active = i; render(); saveState(); };
    sheetsEl.appendChild(btn);
  });
}

function renderTable(){
  const sheet = state.sheets[state.active];
  if(!sheet) return;
  document.getElementById('sheetTitle').textContent = sheet.name;
  document.getElementById('sheetPriority').textContent = 'Priority: ' + sheet.priority;
  document.getElementById('sheetComments').textContent = sheet.comments || '—';
  document.getElementById('prioritySelect').value = sheet.priority;
  document.getElementById('commentsInput').value = sheet.comments || '';

  const rows = sheet.data.length; const cols = sheet.data[0]?.length || 0;
  tableEl.innerHTML = '';
  const thead = document.createElement('thead');
  const headRow = document.createElement('tr');
  headRow.appendChild(document.createElement('th'));
  for(let c=0;c<cols;c++){
    const th = document.createElement('th'); th.textContent = colLabel(c);
    headRow.appendChild(th);
  }
  thead.appendChild(headRow); tableEl.appendChild(thead);

  const tbody = document.createElement('tbody');
  for(let r=0;r<rows;r++){
    const tr = document.createElement('tr');
    const rowHead = document.createElement('th'); rowHead.textContent = r+1; tr.appendChild(rowHead);
    for(let c=0;c<cols;c++){
      const td = document.createElement('td');
      td.contentEditable = true;
      td.dataset.r = r; td.dataset.c = c;
      td.textContent = sheet.data[r][c] || '';
      td.addEventListener('input', onCellInput);
      td.addEventListener('keydown', onCellKeyDown);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  tableEl.appendChild(tbody);
}

/* helpers */
function colLabel(n){ let s=''; n++; while(n>0){ let rem=(n-1)%26; s = String.fromCharCode(65+rem)+s; n = Math.floor((n-1)/26); } return s; }
function onCellInput(e){ const td = e.target; const r = +td.dataset.r; const c = +td.dataset.c; state.sheets[state.active].data[r][c] = td.textContent; saveState(); }
function onCellKeyDown(e){ if(e.key==='Tab'){ e.preventDefault(); const td=e.target; const r=+td.dataset.r, c=+td.dataset.c; focusCell(r,c+1); } }
function focusCell(r,c){ const sel = tableEl.querySelector('td[data-r=\"'+r+'\"][data-c=\"'+c+'\"]'); if(sel){ sel.focus(); placeCaretAtEnd(sel);} }
function placeCaretAtEnd(el){ const range = document.createRange(); const sel = window.getSelection(); range.selectNodeContents(el); range.collapse(false); sel.removeAllRanges(); sel.addRange(range); }

/* table ops */
function addRow(){ const s = state.sheets[state.active]; const cols = s.data[0]?.length || 1; s.data.push(new Array(cols).fill('')); render(); saveState(); }
function addCol(){ const s = state.sheets[state.active]; s.data.forEach(r=>r.push('')); render(); saveState(); }
function delRow(){ const s = state.sheets[state.active]; if(s.data.length<=1) return; s.data.pop(); render(); saveState(); }
function delCol(){ const s = state.sheets[state.active]; if(s.data[0].length<=1) return; s.data.forEach(r=>r.pop()); render(); saveState(); }
function newSheet(){ const name = prompt('New sheet name','Sheet'+(state.sheets.length+1))||('Sheet'+(state.sheets.length+1)); state.sheets.push(defaultSheet(name)); state.active = state.sheets.length-1; render(); saveState(); }
function renameSheet(){ const name = prompt('Rename sheet', state.sheets[state.active].name)||state.sheets[state.active].name; state.sheets[state.active].name = name; render(); saveState(); }
function deleteSheet(){ if(state.sheets.length===1){ if(!confirm('Delete last sheet? This will clear it.')) return; state.sheets[0]=defaultSheet('Sheet1'); state.active=0; render(); saveState(); return; } if(!confirm('Delete this sheet?')) return; state.sheets.splice(state.active,1); state.active = Math.max(0,state.active-1); render(); saveState(); }
function clearSheet(){ if(!confirm('Clear sheet?')) return; const name = state.sheets[state.active].name; state.sheets[state.active] = defaultSheet(name,10,6); render(); saveState(); }

/* find */
document.getElementById('findInput').addEventListener('keydown', (e)=>{ if(e.key==='Enter'){ const q=e.target.value.trim(); if(!q) return; const s=state.sheets[state.active]; for(let r=0;r<s.data.length;r++){ for(let c=0;c<s.data[0].length;c++){ if((s.data[r][c]||'').toString().includes(q)){ focusCell(r,c); return; } } } alert('Not found'); } });

/* priorities & comments */
document.getElementById('prioritySelect').addEventListener('change', (e)=>{ state.sheets[state.active].priority = e.target.value; render(); saveState(); });
document.getElementById('commentsInput').addEventListener('input', (e)=>{ state.sheets[state.active].comments = e.target.value; saveState(); });

/* import/export local */
function exportJSON(){ const data = JSON.stringify(state); const blob = new Blob([data],{type:'application/json'}); downloadBlob(blob,'local-excel-state.json'); }
function importJSON(){ const input = document.createElement('input'); input.type='file'; input.accept='application/json'; input.onchange = ()=>{ const f = input.files[0]; if(!f) return; const reader = new FileReader(); reader.onload = ()=>{ try{ const parsed = JSON.parse(reader.result); if(parsed.sheets){ state = parsed; render(); saveState(); alert('Imported'); }else alert('Invalid file'); }catch(e){ alert('Invalid JSON'); }}; reader.readAsText(f); }; input.click(); }
function exportCSV(){ const s = state.sheets[state.active]; const rows = s.data.map(r => r.map(c=>escapeCSV(c)).join(',')); const csv = rows.join('\\n'); downloadBlob(new Blob([csv],{type:'text/csv'}), s.name.replace(/\\s+/g,'_') + '.csv'); }
document.getElementById('importCSV').addEventListener('change', (e)=>{ const f = e.target.files[0]; if(!f) return; const reader = new FileReader(); reader.onload = ()=>{ const parsed = simpleParseCSV(reader.result); state.sheets[state.active].data = parsed; render(); saveState(); alert('CSV imported'); }; reader.readAsText(f); });
function simpleParseCSV(text){ const rows = text.split(/\\r?\\n/).filter(r=>r.length>0); return rows.map(r=> r.split(',')); }
function escapeCSV(val){ if(val==null) return ''; val = val.toString(); if(val.includes(',')||val.includes('\\n')||val.includes('"')) return '\"'+val.replace(/\"/g,'\"\"')+'\"'; return val; }
function downloadBlob(blob,name){ const url=URL.createObjectURL(blob); const a=document.createElement('a'); a.href=url; a.download=name; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url); }

/* GitHub API helpers
   - User provides token (input), owner, repo, path prefix.
   - Save sheet as JSON file at path: prefix + encodeURIComponent(sheetName) + '.json'
   - If file exists, fetch SHA then PUT to update. If not, PUT without sha to create.
*/
async function ghFetch(path, token, method='GET', body=null){
  const url = `https://api.github.com/repos/${path}`;
  const opts = { method, headers: { Accept:'application/vnd.github.v3+json' } };
  if(token) opts.headers['Authorization'] = 'token ' + token;
  if(body) { opts.body = JSON.stringify(body); opts.headers['Content-Type']='application/json'; }
  const res = await fetch(url, opts);
  return res;
}

async function getFileSha(owner, repo, fullpath, token){
  const endpoint = `${owner}/${repo}/contents/${encodeURIComponent(fullpath)}`;
  const res = await ghFetch(endpoint, token, 'GET');
  if(res.status===200){ const j = await res.json(); return j.sha; }
  return null;
}

async function putFile(owner, repo, fullpath, contentStr, message, token, sha=null){
  const endpoint = `${owner}/${repo}/contents/${encodeURIComponent(fullpath)}`;
  const body = {
    message: message || 'Update sheet ' + fullpath,
    content: b64EncodeUnicode(contentStr),
    committer: { name: 'web-app', email: 'noreply@example.com' }
  };
  if(sha) body.sha = sha;
  const res = await ghFetch(endpoint, token, 'PUT', body);
  return res;
}

function b64EncodeUnicode(str) {
  // base64 encode unicode
  return btoa(unescape(encodeURIComponent(str)));
}

/* Save single sheet to GitHub */
async function saveActiveToGitHub(){
  const token = document.getElementById('ghToken').value.trim();
  const owner = document.getElementById('ghOwner').value.trim();
  const repo = document.getElementById('ghRepo').value.trim();
  const prefix = document.getElementById('ghPath').value.trim();
  if(!token || !owner || !repo){ alert('Provide token, owner and repo'); return; }
  const sheet = state.sheets[state.active];
  const filename = (prefix ? prefix.replace(/\\/$/, '') + '/' : '') + sanitizeFileName(sheet.name) + '.json';
  document.getElementById('saveStatus').textContent = 'Saving to GitHub...';
  try{
    const sha = await getFileSha(owner, repo, filename, token);
    const contentStr = JSON.stringify(sheet, null, 2);
    const message = `Save sheet ${sheet.name} (via web app)`;
    const res = await putFile(owner, repo, filename, contentStr, message, token, sha);
    if(res.status===201 || res.status===200){
      document.getElementById('saveStatus').textContent = 'Saved to GitHub: ' + filename;
      alert('Saved to GitHub: ' + filename);
    }else{
      const j = await res.json(); console.error(j);
      alert('GitHub error: ' + (j.message || res.status));
      document.getElementById('saveStatus').textContent = 'GitHub save failed';
    }
  }catch(e){ console.error(e); alert('Error: ' + e.message); document.getElementById('saveStatus').textContent = 'GitHub save failed'; }
}

function sanitizeFileName(name){
  return name.replace(/[\\/?%*:|\"<>]/g, '-').replace(/\\s+/g, '_');
}

/* Save all sheets */
async function saveAllToGitHub(){
  const token = document.getElementById('ghToken').value.trim();
  const owner = document.getElementById('ghOwner').value.trim();
  const repo = document.getElementById('ghRepo').value.trim();
  const prefix = document.getElementById('ghPath').value.trim();
  if(!token || !owner || !repo){ alert('Provide token, owner and repo'); return; }
  document.getElementById('saveStatus').textContent = 'Saving all to GitHub...';
  for(let i=0;i<state.sheets.length;i++){
    const sheet = state.sheets[i];
    const filename = (prefix ? prefix.replace(/\\/$/, '') + '/' : '') + sanitizeFileName(sheet.name) + '.json';
    try{
      const sha = await getFileSha(owner, repo, filename, token);
      const res = await putFile(owner, repo, filename, JSON.stringify(sheet, null, 2), `Save sheet ${sheet.name}`, token, sha);
      if(!(res.status===201 || res.status===200)) { const j=await res.json(); console.error('error saving', sheet.name, j); }
    }catch(e){ console.error('save error', e); }
  }
  document.getElementById('saveStatus').textContent = 'Saved all to GitHub';
  alert('Saved all sheets (attempted). Check repo.');
}

/* wiring */
loadState(); render();
function render(){ renderSheetsList(); renderTable(); }

document.getElementById('addRow').onclick = addRow;
document.getElementById('addCol').onclick = addCol;
document.getElementById('delRow').onclick = delRow;
document.getElementById('delCol').onclick = delCol;
document.getElementById('newSheet').onclick = newSheet;
document.getElementById('renameSheet').onclick = renameSheet;
document.getElementById('deleteSheet').onclick = deleteSheet;
document.getElementById('clearSheet').onclick = clearSheet;
document.getElementById('exportAll').onclick = exportJSON;
document.getElementById('importAll').onclick = importJSON;
document.getElementById('exportCSV').onclick = exportCSV;
document.getElementById('saveToGH').onclick = saveActiveToGitHub;
document.getElementById('saveAllToGH').onclick = saveAllToGitHub;

/* simple UI: show file name when clicking a sheet */
document.getElementById('importAll').addEventListener('click', importJSON);

/* keyboard ctrl+s to save to local */
window.addEventListener('keydown', (e)=>{ if((e.ctrlKey||e.metaKey) && e.key.toLowerCase()==='s'){ e.preventDefault(); saveState(); alert('Saved locally'); } });

/* initial local save stamp */
saveState();

/* helpers for import JSON */
function importJSON(){
  const input = document.createElement('input'); input.type='file'; input.accept='application/json';
  input.onchange = ()=>{ const f = input.files[0]; if(!f) return; const r = new FileReader(); r.onload = ()=>{ try{ const parsed = JSON.parse(r.result); if(parsed.sheets){ state = parsed; render(); saveState(); alert('Imported'); } else alert('Invalid file'); }catch(e){ alert('Invalid JSON'); }}; r.readAsText(f); }; input.click();
}
</script>

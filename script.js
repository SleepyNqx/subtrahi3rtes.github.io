// Simple local Excel-like app.
// Persists to localStorage under key 'localExcel.sheets'

const STORAGE_KEY = 'localExcel.sheets.v1';
let state = { sheets: [], active: 0 };
let saveTimer = null;
const statusEl = document.getElementById('status');
const sheetsEl = document.getElementById('sheets');
const tableEl = document.getElementById('sheetTable');
const autosaveSel = document.getElementById('autosave');

function defaultSheet(name='Sheet1', rows=10, cols=6){
  const data = [];
  for(let r=0;r<rows;r++){ data.push(new Array(cols).fill('')); }
  return { name, data };
}

function loadState(){
  try{
    const raw = localStorage.getItem(STORAGE_KEY);
    if(raw){ state = JSON.parse(raw); }
    if(!state.sheets || state.sheets.length===0){ state = { sheets:[defaultSheet('Sheet1')], active:0 }; }
  }catch(e){ console.error(e); state = { sheets:[defaultSheet('Sheet1')], active:0 }; }
}

function saveState(immediate=false){
  if(autosaveSel.value==='off' && !immediate) return;
  clearTimeout(saveTimer);
  if(immediate){ localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); statusEl.textContent='Saved'; return; }
  statusEl.textContent='Saving...';
  saveTimer = setTimeout(()=>{ localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); statusEl.textContent='Saved'; }, 350);
}

function renderSheetsList(){
  sheetsEl.innerHTML='';
  state.sheets.forEach((s,i)=>{
    const btn = document.createElement('button');
    btn.className='sheet-btn' + (i===state.active ? ' active':'');
    btn.textContent = s.name || ('Sheet '+(i+1));
    btn.onclick = ()=>{ state.active = i; render(); saveState(); };
    sheetsEl.appendChild(btn);
  });
}

function renderTable(){
  const sheet = state.sheets[state.active];
  if(!sheet) return;
  const rows = sheet.data.length; const cols = sheet.data[0]?.length || 0;
  tableEl.innerHTML='';
  const thead = document.createElement('thead');
  const headRow = document.createElement('tr');
  const corner = document.createElement('th'); corner.textContent=''; headRow.appendChild(corner);
  for(let c=0;c<cols;c++){ const th = document.createElement('th'); th.textContent = colLabel(c); headRow.appendChild(th); }
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

function colLabel(n){
  let s=''; n++; while(n>0){ let rem=(n-1)%26; s = String.fromCharCode(65+rem)+s; n = Math.floor((n-1)/26); } return s;
}

function onCellInput(e){
  const td = e.target; const r = +td.dataset.r; const c = +td.dataset.c;
  state.sheets[state.active].data[r][c] = td.textContent;
  saveState();
}

function onCellKeyDown(e){
  if(e.key==='Tab'){ e.preventDefault(); const td = e.target; const r=+td.dataset.r, c=+td.dataset.c; focusCell(r, c+1); }
}

function focusCell(r,c){
  const selector = 'td[data-r=\"'+r+'\"][data-c=\"'+c+'\"]';
  const el = tableEl.querySelector(selector);
  if(el){ el.focus(); placeCaretAtEnd(el); }
}

function placeCaretAtEnd(el){
  const range = document.createRange(); const sel = window.getSelection(); range.selectNodeContents(el); range.collapse(false); sel.removeAllRanges(); sel.addRange(range);
}

function addRow(){
  const sheet = state.sheets[state.active];
  const cols = sheet.data[0]?.length || 1; sheet.data.push(new Array(cols).fill(''));
  render(); saveState();
}
function addCol(){
  const sheet = state.sheets[state.active];
  sheet.data.forEach(row=>row.push(''));
  render(); saveState();
}
function delRow(){
  const sheet = state.sheets[state.active]; if(sheet.data.length<=1) return; sheet.data.pop(); render(); saveState();
}
function delCol(){
  const sheet = state.sheets[state.active]; if(sheet.data[0].length<=1) return; sheet.data.forEach(row=>row.pop()); render(); saveState();
}

function newSheet(){
  const name = prompt('New sheet name','Sheet'+(state.sheets.length+1))||('Sheet'+(state.sheets.length+1));
  state.sheets.push(defaultSheet(name)); state.active = state.sheets.length-1; render(); saveState(true);
}
function renameSheet(){
  const name = prompt('Rename sheet', state.sheets[state.active].name)||state.sheets[state.active].name; state.sheets[state.active].name = name; renderSheetsList(); saveState();
}
function deleteSheet(){
  if(state.sheets.length===1){ if(!confirm('Delete last sheet? This will clear it.')) return; state.sheets[0]=defaultSheet('Sheet1'); state.active=0; render(); saveState(true); return; }
  if(!confirm('Delete this sheet?')) return;
  state.sheets.splice(state.active,1); state.active = Math.max(0,state.active-1); render(); saveState(true);
}

function clearSheet(){ if(!confirm('Clear sheet?')) return; const name = state.sheets[state.active].name; state.sheets[state.active]=defaultSheet(name,10,6); render(); saveState(true); }

function exportJSON(){ const data = JSON.stringify(state); const blob = new Blob([data],{type:'application/json'}); downloadBlob(blob,'local-excel.json'); }
function importJSON(){ const input = document.createElement('input'); input.type='file'; input.accept='application/json'; input.onchange = ()=>{
  const f = input.files[0]; if(!f) return; const reader = new FileReader(); reader.onload = ()=>{ try{ const parsed = JSON.parse(reader.result); if(parsed.sheets) state = parsed; render(); saveState(true); alert('Imported'); }catch(e){ alert('Invalid file'); }}; reader.readAsText(f);
}; input.click(); }

function exportCSV(){
  const sheet = state.sheets[state.active]; const rows = sheet.data.map(row=> row.map(cell=> escapeCSV(cell)).join(','));
  const csv = rows.join('\n'); const blob = new Blob([csv],{type:'text/csv'}); downloadBlob(blob,sheet.name.replace(/\s+/g,'_')+'.csv');
}
function importCSVFile(file){ const reader = new FileReader(); reader.onload = ()=>{ const text = reader.result; const parsed = simpleParseCSV(text); state.sheets[state.active].data = parsed; render(); saveState(true); alert('CSV imported'); }; reader.readAsText(file); }

function simpleParseCSV(text){
  const rows = text.split(/\r?\n/).filter(r=>r.length>0);
  return rows.map(r=> r.split(','));
}
function escapeCSV(val){ if(val==null) return ''; if(val.includes(',')||val.includes('\n')||val.includes('"')) return '"'+val.replace(/"/g,'""')+'"'; return val; }

function downloadBlob(blob, name){ const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url); }

function findInSheet(query){
  query = query.trim(); if(!query) return;
  const sheet = state.sheets[state.active];
  for(let r=0;r<sheet.data.length;r++){
    for(let c=0;c<sheet.data[0].length;c++){
      if((sheet.data[r][c]||'').toString().includes(query)){ focusCell(r,c); return; }
    }
  }
  alert('Not found');
}

function render(){ renderSheetsList(); renderTable(); }

// wire UI
loadState(); render();

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
document.getElementById('importCSV').addEventListener('change', (e)=>{ const f = e.target.files[0]; if(f) importCSVFile(f); });

document.getElementById('findInput').addEventListener('keydown', (e)=>{ if(e.key==='Enter'){ findInSheet(e.target.value); } });

autosaveSel.addEventListener('change', ()=>{ saveState(true); });

// keyboard shortcuts: ctrl+s save
window.addEventListener('keydown', (e)=>{ if((e.ctrlKey||e.metaKey) && e.key.toLowerCase()==='s'){ e.preventDefault(); saveState(true); } });

// initial save state
saveState(true);

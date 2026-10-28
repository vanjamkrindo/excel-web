/* ============ 01 · Model & State (multi-workbook) ============ */
const NROWS = 200, NCOLS = 30; // A..AD
const $ = s => document.querySelector(s);
const $$ = s => [...document.querySelectorAll(s)];
const esc = s => String(s ?? "").replace(/[&<>\"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));

function colName(n){ let s=""; n++; while(n>0){ n--; s=String.fromCharCode(65+n%26)+s; n=Math.floor(n/26);} return s; }
function colIdx(s){ let n=0; for(const c of s.toUpperCase()) n=n*26+(c.charCodeAt(0)-64); return n-1; }
function cellRef(r,c){ return colName(c)+ (r+1); }
function parseRef(ref){
  const m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(ref.trim()); if(!m) return null;
  return {r:parseInt(m[2],10)-1, c:colIdx(m[1])};
}
function clampRC(r,c){ return {r:Math.max(0,Math.min(NROWS-1,r)), c:Math.max(0,Math.min(NCOLS-1,c))}; }
function rangeOf(a,b){ return {r1:Math.min(a.r,b.r), c1:Math.min(a.c,b.c), r2:Math.max(a.r,b.r), c2:Math.max(a.c,b.c)}; }

/* Workbook: sheets[] {name, cells:{"r,c":{v, f, s:{...style}, numfmt, m:mergeId}}, cols:{w}, rows:{h}, merges:[{r1,c1,r2,c2}], frozen:{rows,cols}, charts:[] } */

/* ---------- multi-workbook model ----------
   WORKBOOKS = [{id, name, sheets, active, cur:{sheet,r,c}, dirty, savedTs}]
   ACTIVE    = index of the workbook currently focused
   `wb` / `cur` stay as live aliases so every existing function keeps working.
------------------------------------------- */
let WORKBOOKS = [];
let ACTIVE = 0;
let NEXT_ID = 1;

let wb = null;
let cur = {sheet:0, r:0, c:0};           // active cell (alias of active workbook)
let selAnchor = null;                     // selection anchor while extending
let editing = false, editFromFx = false;
let undoStack = [], redoStack = [];
let clipboard = null; // {cells:{key:{v,f,s,numfmt}}, w,h, cut}
let dirty = false;

function activeWB(){ return WORKBOOKS[ACTIVE] || WORKBOOKS[0] || null; }
function wbIndexById(id){ return WORKBOOKS.findIndex(w=>w.id===id); }
function wbById(id){ const i = wbIndexById(id); return i<0 ? null : WORKBOOKS[i]; }

/* Point the module-level wb/cur aliases at the active workbook. */
function syncAliases(){
  const w = activeWB();
  if(!w){ wb = null; return; }
  wb = w;
  cur = w.cur || (w.cur = {sheet:0, r:0, c:0});
  dirty = !!w.dirty;
}
/* Write alias state back into the workbook object (call before saving). */
function commitAliases(){
  const w = activeWB(); if(!w) return;
  w.cur = cur;
  w.active = cur.sheet;
  w.dirty = dirty;
}

function newWBID(){ return NEXT_ID++; }

/* Create a workbook and register it. Returns the workbook object. */
function createWorkbook(name, opts){
  const w = {
    id: newWBID(),
    name: name || nextBookName(),
    sheets: [blankSheet("Sheet1")],
    active: 0,
    cur: {sheet:0, r:0, c:0},
    dirty: false,
    savedTs: 0,
    path: (opts && opts.path) || null,
  };
  WORKBOOKS.push(w);
  return w;
}
function nextBookName(){
  const used = new Set(WORKBOOKS.map(w=>w.name.toLowerCase()));
  for(let i=1;i<1000;i++){ const n = "Book"+i; if(!used.has(n.toLowerCase())) return n; }
  return "Book"+Date.now();
}
/* Keep a name unique among the other workbooks (used on open/rename). */
function uniqueName(name, exceptId){
  const base = String(name||"Book1").replace(/\.(xlsx|xls|csv)$/i,"");
  let nm = base, i = 1;
  const used = new Set(WORKBOOKS.filter(w=>w.id!==exceptId).map(w=>w.name.toLowerCase()));
  while(used.has(nm.toLowerCase())){ nm = base + " (" + (++i) + ")"; }
  return nm;
}
function switchWB(i){
  if(i<0 || i>=WORKBOOKS.length || i===ACTIVE) return;
  commitAliases();
  ACTIVE = i;
  syncAliases();
  undoStack.length = 0; redoStack.length = 0;   // history is not shared across workbooks
  selAnchor = null; editing = false;
  clampCur();
  renderAll();
  afterChange();
}
function closeWB(i){
  if(WORKBOOKS.length<=1){ flashStatus("At least one workbook must stay open"); return; }
  const w = WORKBOOKS[i]; if(!w) return;
  if(w.dirty && !confirm(`"${w.name}" has unsaved changes. Close anyway?`)) return;
  WORKBOOKS.splice(i,1);
  if(ACTIVE>=WORKBOOKS.length) ACTIVE = WORKBOOKS.length-1;
  syncAliases();
  renderAll(); afterChange();
}
function renderAll(){
  renderWorkbookBar();
  setDocTitle();
  renderSheets();
  renderGrid();
}

function sheet(){ return wb.sheets[cur.sheet]; }
function ck(r,c){ return r+","+c; }
function getCell(r,c){ return sheet().cells[ck(r,c)] || null; }
function getCellW(r,c){ const s=wb.sheets[cur.sheet]; return s.cells[ck(r,c)] || (s.cells[ck(r,c)]={}); }
function getStyle(r,c){ const cl=getCell(r,c); return cl && cl.s ? cl.s : {}; }
function numFmtOf(r,c){ const cl=getCell(r,c); return cl && cl.numfmt ? cl.numfmt : "General"; }

function snapshot(){
  return JSON.stringify(wb.sheets[cur.sheet]);
}
function pushUndo(label){
  undoStack.push({sheet:cur.sheet, data:snapshot(), label});
  if(undoStack.length>100) undoStack.shift();
  redoStack.length=0;
}
function pushUndoAll(){
  undoStack.push({sheet:-1, data:JSON.stringify(wb.sheets), label:"all sheets"});
  redoStack.length=0;
}
function undo(){
  const u = undoStack.pop(); if(!u) return;
  redoStack.push({sheet:u.sheet, data: u.sheet<0 ? JSON.stringify(wb.sheets) : snapshot(), label:u.label});
  if(u.sheet<0){ wb.sheets = JSON.parse(u.data); if(cur.sheet>=wb.sheets.length) cur.sheet=0; renderSheets(); }
  else wb.sheets[u.sheet] = JSON.parse(u.data);
  clampCur(); renderGrid(); afterChange();
}
function redo(){
  const u = redoStack.pop(); if(!u) return;
  undoStack.push(u);
  if(u.sheet<0){ wb.sheets = JSON.parse(u.data); renderSheets(); }
  else wb.sheets[u.sheet] = JSON.parse(u.data);
  clampCur(); renderGrid(); afterChange();
}
function clampCur(){ cur = {...cur, ...clampRC(cur.r, cur.c)}; }

function newWorkbook(){
  return { name:"Book1", sheets:[ blankSheet("Sheet1") ], active:0 };
}
function markDirty(){ dirty=true; scheduleSave(); }

/* selection helpers */
function selRange(){
  return selAnchor ? rangeOf(selAnchor, cur) : {r1:cur.r,c1:cur.c,r2:cur.r,c2:cur.c};
}
function eachSelCell(fn){
  const s = selRange();
  for(let r=s.r1;r<=s.r2;r++) for(let c=s.c1;c<=s.c2;c++) fn(r,c);
}
function fmtNum(v){
  if(typeof v === "number"){
    if(Number.isInteger(v)) return String(v);
    const a = Math.abs(v);
    if(a>=1e11 || (a<1e-9 && a>0)) return v.toExponential(4);
    return String(Math.round(v*1e10)/1e10);
  }
  return String(v);
}

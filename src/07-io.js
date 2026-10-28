/* ============ 07 · Sheets, File I/O, Keyboard, Init ============ */
const LS_KEY = "excel-web-workbook-v1";

/* ---------- document title ---------- */
function setDocTitle(){
  const nm = (wb && wb.name) ? wb.name : "Book1";
  const el = $("#doctitle"); if(el) el.textContent = nm;   // Excel shows the bare workbook name in the title bar
  document.title = nm + " - Excel";
}

/* ---------- sheet tabs ---------- */
function renderSheets(){
  const bar = $("#sheettabs");
  const visible = wb.sheets.map((s,i)=>({s,i})).filter(x=>!x.s.hidden);
  bar.innerHTML =
    `<button class="navbtn" id="sh-first" title="First Sheet">⏮</button>` +
    `<button class="navbtn" id="sh-prev" title="Previous Sheet">◀</button>` +
    `<button class="navbtn" id="sh-next" title="Next Sheet">▶</button>` +
    visible.map(({s,i})=>`<button class="shbtn${i===cur.sheet?" active":""}" data-sh="${i}"${s.tabColor?` style="box-shadow:inset 0 -3px 0 ${s.tabColor}"`:""}>${esc(s.name)}</button>`).join("") +
    `<button class="shbtn" id="sheetadd" title="New sheet">＋</button>`;
  syncChecks();
}
document.addEventListener("click", e=>{
  const sh = e.target.closest("[data-sh]");
  if(sh){ switchSheet(+sh.dataset.sh); return; }
  if(e.target.id==="sheetadd"){ addSheet(); return; }
  if(e.target.id==="sh-prev"){ stepSheet(-1); return; }
  if(e.target.id==="sh-next"){ stepSheet(1); return; }
  if(e.target.id==="sh-first"){ switchSheet(0); return; }
});
document.addEventListener("dblclick", e=>{
  const sh = e.target.closest("[data-sh]");
  if(sh) renameSheet(+sh.dataset.sh);
});
document.addEventListener("contextmenu", e=>{
  const sh = e.target.closest("[data-sh]");
  if(sh){ e.preventDefault(); sheetMenu(+sh.dataset.sh, e.clientX, e.clientY); }
});
function stepSheet(d){
  const vis = wb.sheets.map((s,i)=>({s,i})).filter(x=>!x.s.hidden).map(x=>x.i);
  const at = vis.indexOf(cur.sheet);
  const nx = vis[Math.max(0, Math.min(vis.length-1, at+d))];
  if(nx!==undefined) switchSheet(nx);
}
function switchSheet(i){
  if(i===cur.sheet) return;
  cur.sheet = i; cur.r=0; cur.c=0; selAnchor=null;
  recalcAll(); renderGrid(); drawAllCharts();
  positionSel(); syncToolbar(); applyPageSetup();
}
function addSheet(){
  pushUndoAll();
  let n = wb.sheets.length+1;
  while(wb.sheets.some(s=>s.name==="Sheet"+n)) n++;
  wb.sheets.push(blankSheet("Sheet"+n));
  cur.sheet = wb.sheets.length-1; cur.r=0; cur.c=0;
  recalcAll(); renderGrid(); markDirty();
}
function renameSheet(i){
  const s = wb.sheets[i];
  dlg("dlg-ren","Rename Sheet",`<div class="row"><label>Name</label><input type="text" id="sn" value="${esc(s.name)}"></div>`, d=>{
    const v = d.querySelector("#sn").value.trim();
    if(!v) return false;
    if(wb.sheets.some((x,j)=>j!==i && x.name===v)){ alert("That sheet name is already in use."); return false; }
    pushUndoAll(); s.name = v; renderSheets(); markDirty();
  }, "Rename");
}
function sheetMenu(i,x,y){
  hideCtx();
  const m = document.createElement("div");
  m.className="popover show"; m.style.left=Math.min(x,innerWidth-200)+"px"; m.style.top=Math.max(4,y-120)+"px";
  m.innerHTML = `
  <button class="pitem" data-shact="insert"><span class="pi">＋</span>Insert…</button>
  <button class="pitem" data-shact="rename"><span class="pi">✎</span>Rename</button>
  <button class="pitem" data-shact="dup"><span class="pi">⧉</span>Move or Copy…</button>
  <div class="psep"></div>
  <button class="pitem" data-shact="hide"><span class="pi">👁</span>Hide</button>
  <button class="pitem" data-shact="unhide"><span class="pi">👁</span>Unhide…</button>
  <div class="psep"></div>
  <button class="pitem" data-shact="tabcolor"><span class="pi">🎨</span>Tab Color</button>
  <button class="pitem" data-shact="protect"><span class="pi">🔒</span>Protect Sheet</button>
  <div class="psep"></div>
  <button class="pitem" data-shact="del"><span class="pi">🗑</span>Delete</button>`;
  document.body.appendChild(m); ctxMenu = m;
  m.addEventListener("click", e=>{
    const b = e.target.closest("[data-shact]"); if(!b) return;
    const act = b.dataset.shact; hideCtx();
    if(act==="insert"){ pushUndoAll(); wb.sheets.splice(i+1,0,blankSheet("Sheet"+(wb.sheets.length+1))); renderSheets(); markDirty(); }
    if(act==="rename") renameSheet(i);
    if(act==="dup"){ pushUndoAll(); const copy=JSON.parse(JSON.stringify(wb.sheets[i])); copy.name=wb.sheets[i].name+" (2)"; wb.sheets.splice(i+1,0,copy); renderSheets(); markDirty(); }
    if(act==="hide"){
      const vis = wb.sheets.filter(s=>!s.hidden);
      if(vis.length<=1){ alert("A workbook must contain at least one visible sheet."); return; }
      pushUndoAll(); wb.sheets[i].hidden = true;
      if(cur.sheet===i){ const nx = wb.sheets.findIndex(s=>!s.hidden); cur.sheet = nx<0?0:nx; }
      renderSheets(); recalcAll(); renderGrid(); markDirty();
    }
    if(act==="unhide"){
      const hid = wb.sheets.map((s,j)=>({s,j})).filter(x=>x.s.hidden);
      if(!hid.length){ alert("No hidden sheets in this workbook."); return; }
      dlg("dlg-unhide","Unhide",`<div class="row"><label>Sheet</label><select id="uh">${hid.map(h=>`<option value="${h.j}">${esc(h.s.name)}</option>`).join("")}</select></div>`, d=>{
        pushUndoAll(); wb.sheets[+d.querySelector("#uh").value].hidden = false; renderSheets(); markDirty();
      }, "OK");
    }
    if(act==="tabcolor"){
      const colors = ["#217346","#c00000","#ed7d31","#ffc000","#4472c4","#7030a0","#00b0f0","#a5a5a5"];
      dlg("dlg-tabcolor","Tab Color",`<div style="display:flex;gap:6px;flex-wrap:wrap">${colors.map(c=>`<div class="sw" data-tc="${c}" style="background:${c};width:26px;height:26px"></div>`).join("")}</div>`, ()=>false, "Close");
      const d = $("#dlg-tabcolor");
      d.addEventListener("click", ev=>{
        const sw = ev.target.closest("[data-tc]"); if(!sw) return;
        pushUndoAll(); wb.sheets[i].tabColor = sw.dataset.tc; renderSheets(); markDirty(); d.close();
      });
    }
    if(act==="protect"){ pushUndoAll(); wb.sheets[i].protect = !wb.sheets[i].protect; flashStatus(wb.sheets[i].protect?"Sheet protected":"Protection removed"); renderGrid(); markDirty(); }
    if(act==="del"){
      if(wb.sheets.length===1){ alert("A workbook must contain at least one sheet."); return; }
      if(!confirm("Delete sheet \""+wb.sheets[i].name+"\"?")) return;
      pushUndoAll(); wb.sheets.splice(i,1);
      if(cur.sheet>=wb.sheets.length) cur.sheet=wb.sheets.length-1;
      recalcAll(); renderGrid(); markDirty();
    }
  });
}

/* ---------- after change ---------- */
function afterChange(){
  recalcAll(); renderGrid(); drawAllCharts(); positionSel(); syncToolbar(); markDirty();
  if(typeof renderSplitGrids === "function") renderSplitGrids();
  if(typeof renderWorkbookBar === "function") renderWorkbookBar();
}
let saveTimer = null;
function scheduleSave(){
  dirty = true;
  const a = $("#autosave");
  if(typeof AUTOSAVE_ON !== "undefined" && !AUTOSAVE_ON){
    if(a) a.textContent = "Unsaved changes";
    return;
  }
  if(a) a.textContent = "Saving…";
  clearTimeout(saveTimer);
  saveTimer = setTimeout(saveLocal, 700);
}
/* ---------- multi-workbook persistence ---------- */
const LS_ALL = "excel-web-workbooks-v1";
const LS_ACTIVE = "excel-web-active-v1";

function saveLocalAll(){
  try{
    commitAliases();
    const payload = WORKBOOKS.map(w=>({
      id:w.id, name:w.name, sheets:w.sheets, active:w.active||0,
      cur:w.cur||{sheet:0,r:0,c:0}, savedTs:w.savedTs||0, path:w.path||null
    }));
    localStorage.setItem(LS_ALL, JSON.stringify({v:1, books:payload, next:NEXT_ID}));
    localStorage.setItem(LS_ACTIVE, String(ACTIVE));
    commitAliases();
    WORKBOOKS.forEach(w=>{ w.dirty = false; });
    dirty = false;
    const a = $("#autosave"); if(a) a.textContent = "Saved to this device";
  }catch(e){ const a = $("#autosave"); if(a) a.textContent = "Not saved"; }
}
function saveLocal(){
  const a = $("#autosave");
  if(typeof AUTOSAVE_ON !== "undefined" && !AUTOSAVE_ON){
    if(a) a.textContent = "Unsaved changes";
    return;
  }
  if(a) a.textContent = "Saving…";
  clearTimeout(saveTimer);
  saveTimer = setTimeout(saveLocalAll, 700);
}
function loadLocalAll(){
  try{
    const raw = localStorage.getItem(LS_ALL);
    if(!raw) return null;
    const d = JSON.parse(raw);
    if(!d || !Array.isArray(d.books) || !d.books.length) return null;
    return d;
  }catch(e){ return null; }
}
/* Legacy single-workbook slot, migrated on first run. */
function loadLocal(){
  try{
    const raw = localStorage.getItem(LS_KEY);
    if(!raw) return null;
    const d = JSON.parse(raw);
    if(!d || !d.wb || !d.wb.sheets) return null;
    return d;
  }catch(e){ return null; }
}

/* ---------- xlsx / csv ---------- */
function cellStyleForXlsx(st, numfmt){
  const s = {};
  if(st.bold||st.italic||st.underline||st.color||st.size||st.font) s.font = {
    bold:!!st.bold, italic:!!st.italic, underline:!!st.underline,
    color:{rgb:(st.color||"#000000").replace("#","FF").toUpperCase()},
    sz: st.size||11, name: st.font||"Calibri"
  };
  if(st.bg||st.align||st.valign||st.wrap) s.alignment = {
    horizontal: st.align||undefined, vertical: st.valign||undefined, wrapText: !!st.wrap
  };
  if(st.bg) s.fill = {fgColor:{rgb:st.bg.replace("#","FF").toUpperCase()}, patternType:"solid"};
  const nf = {General:"General", Number:"#,##0.00", Currency:'$#,##0.00', Accounting:'$#,##0.00',
    Percent:"0.00%", ShortDate:"mm/dd/yyyy", LongDate:"dddd, mmmm dd, yyyy", Time:"hh:mm:ss", Text:"@"}[numfmt];
  if(nf) s.numFmt = nf;
  return (Object.keys(s).length ? s : null);
}
function buildWorkbookObject(){
  const wbo = XLSX.utils.book_new();
  wb.sheets.forEach(sh=>{
    const ws = {};
    let maxR=0, maxC=0;
    for(const k in sh.cells){
      const [r,c] = k.split(",").map(Number);
      const cl = sh.cells[k];
      const addr = XLSX.utils.encode_cell({r,c});
      maxR = Math.max(maxR,r); maxC = Math.max(maxC,c);
      if(cl.f){
        const f = cl.f.replace(/^=/,"");
        const v = evalCache[k];
        const cell = {f};
        if(typeof v==="number"){ cell.v = v; cell.t="n"; }
        else if(v!=="" && v!=null){ cell.v = String(v); cell.t="s"; }
        ws[addr] = cell;
      } else if(cl.v!=null && cl.v!==""){
        ws[addr] = typeof cl.v==="number" ? {v:cl.v, t:"n"} : {v:String(cl.v), t:"s"};
      }
      if(cl.s) ws[addr] = {...(ws[addr]||{t:typeof cl.v==="number"?"n":"s", v:cl.v}), s:cellStyleForXlsx(cl.s, cl.numfmt)};
      if(ws[addr] && !ws[addr].v && !ws[addr].f) ws[addr].v = cl.v ?? "";
    }
    ws["!ref"] = XLSX.utils.encode_range({s:{r:0,c:0}, e:{r:Math.max(maxR,0), c:Math.max(maxC,0)}});
    if(sh.merges.length) ws["!merges"] = sh.merges.map(m=>({s:{r:m.r1,c:m.c1}, e:{r:m.r2,c:m.c2}}));
    const cols = [];
    for(let c=0;c<=maxC;c++) cols.push({wch: Math.max(8, Math.round(colW(c)/7))});
    ws["!cols"] = cols;
    XLSX.utils.book_append_sheet(wbo, ws, sh.name.slice(0,31));
  });
  return wbo;
}
function downloadXlsx(){
  try{
    const wbo = buildWorkbookObject();
    const out = XLSX.write(wbo, {bookType:"xlsx", type:"array", cellStyles:true});
    const blob = new Blob([out], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = (wb.name||"Book1")+".xlsx";
    a.click();
    setTimeout(()=>URL.revokeObjectURL(a.href), 4000);
    flashStatus("Saved as .xlsx");
  }catch(e){ alert("Could not build the .xlsx file: "+e.message); }
}
function saveBook(){ downloadXlsx(); saveLocal(); }
function exportCsv(){
  const sh = sheet();
  const lines = [];
  for(let r=0;r<NROWS;r++){
    const row = [];
    let any = false;
    for(let c=0;c<NCOLS;c++){ const v = displayValue(r,c); if(v!=="") any=true; row.push(/[",;\n]/.test(v)? '"'+v.replace(/"/g,'""')+'"' : v); }
    if(any) lines.push(row.join(","));
  }
  const blob = new Blob(["\ufeff"+lines.join("\r\n")], {type:"text/csv;charset=utf-8"});
  const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=(wb.name||"Book1")+".csv"; a.click();
}
function openPicker(){
  const inp = document.createElement("input");
  inp.type="file"; inp.accept=".xlsx,.xls,.csv,.ods,.txt";
  inp.onchange = ()=>{ const f = inp.files[0]; if(f) readFile(f); };
  inp.click();
}
function readFile(file){
  const rd = new FileReader();
  rd.onload = ()=>{
    try{
      const data = new Uint8Array(rd.result);
      const wbo = XLSX.read(data, {type:"array", cellFormula:true, cellNF:true, cellDates:false});
      wb = {name:file.name.replace(/\.[^.]+$/,""), sheets:[]};
      wbo.SheetNames.forEach(nm=>{
        const ws = wbo.Sheets[nm];
        const sh = blankSheet(nm.slice(0,31));
        const ref = ws["!ref"] || "A1";
        const range = XLSX.utils.decode_range(ref);
        for(let r=range.s.r;r<=Math.min(range.e.r, NROWS-1);r++){
          for(let c=range.s.c;c<=Math.min(range.e.c, NCOLS-1);c++){
            const addr = XLSX.utils.encode_cell({r,c});
            const cell = ws[addr]; if(!cell) continue;
            const out = {};
            if(cell.f) out.f = "="+cell.f;
            if(cell.v!==undefined && cell.v!==null) out.v = cell.v;
            if(cell.z && cell.z!=="General") out.numfmt = { "#,##0.00":"Number", '$#,##0.00':"Currency", "0.00%":"Percent",
              "mm/dd/yyyy":"ShortDate", "m/d/yy":"ShortDate", "dd/mm/yyyy":"ShortDate", "hh:mm:ss":"Time" }[cell.z] || undefined;
            if(!out.f && out.v===undefined) continue;
            sh.cells[ck(r,c)] = out;
          }
        }
        if(ws["!merges"]) sh.merges = ws["!merges"].map(m=>({r1:m.s.r,c1:m.s.c,r2:m.e.r,c2:m.e.c}));
        if(ws["!cols"]) ws["!cols"].forEach((col,i)=>{ if(col && col.wch) sh.cols[i] = Math.round(col.wch*7); });
        wb.sheets.push(sh);
      });
      if(!wb.sheets.length) wb = newWorkbook();
      cur = {sheet:0,r:0,c:0}; selAnchor=null;
      recalcAll(); renderGrid(); drawAllCharts(); setDocTitle(); markDirty();
      flashStatus("Opened: "+file.name);
    }catch(e){ alert("Could not read the file: "+e.message); }
  };
  rd.readAsArrayBuffer(file);
}
/* drag & drop */
["dragover","drop"].forEach(ev=>document.addEventListener(ev, e=>{
  e.preventDefault();
  if(ev==="drop" && e.dataTransfer.files[0]) readFile(e.dataTransfer.files[0]);
}));

/* ---------- backstage (File) ---------- */
const BS_ITEMS = ["Home","New","Open","Info","Save","Save As","Print","Export","Account","Options","Close"];
let bsSel = "Info";
function openBackstage(item){
  $("#backstage").classList.add("show");
  if(item) bsSel = item;
  $("#bs-nav").innerHTML = BS_ITEMS.map(i=>`<button data-bs="${i}" class="${i===bsSel?"sel":""}">${i}</button>`).join("");
  renderBackstage();
}
function renderBackstage(){
  $$("#bs-nav [data-bs]").forEach(b=>b.classList.toggle("sel", b.dataset.bs===bsSel));
  const c = $("#bs-content");
  const cellCount = sheet() ? Object.keys(sheet().cells).length : 0;
  const name = esc(wb.name||"Book1");
  switch(bsSel){
    case "Home":
      c.innerHTML = `<h2 style="font-size:22px;margin-bottom:16px">Good ${new Date().getHours()<12?"morning":"afternoon"}</h2>
      <p style="color:var(--muted);margin-bottom:18px;font-size:13px">${name} · ${wb.sheets.length} sheet(s) · ${cellCount} filled cells on the active sheet</p>
      <div style="display:flex;gap:10px;flex-wrap:wrap">
        <button class="btn primary" data-bsgo="New">New blank workbook</button>
        <button class="btn" data-bsgo="Open">Open</button>
        <button class="btn" data-bsgo="Save As">Save a Copy</button>
      </div>`;
      break;
    case "Info":
      c.innerHTML = `<h2 style="font-size:20px;margin-bottom:14px">${name}</h2>
      <p style="color:var(--muted);margin-bottom:16px;font-size:13px">${wb.sheets.length} sheet(s) · ${cellCount} filled cells on the active sheet · autosaved in this browser</p>
      <h3 style="font-size:13px;margin:18px 0 8px">Compatibility</h3>
      <p style="font-size:12.5px;color:var(--muted);line-height:1.8">Runs fully offline in a single HTML file.
      Native formats: <b>.xlsx</b> (read &amp; write), <b>.csv</b> (read &amp; write), <b>.xls</b> and <b>.ods</b> (read).
      Cell values, formulas, merges and number formats survive a round trip; cell fill/border styling is written by this app only in the browser view.</p>`;
      break;
    case "New":
      c.innerHTML = `<h2 style="font-size:18px;margin-bottom:12px">New</h2>
        <button class="btn primary" id="bs-new">Blank workbook</button>`;
      $("#bs-new").onclick = ()=>{ if(confirm("Create a new workbook? Unsaved changes are lost.")){ wb=newWorkbook(); cur={sheet:0,r:0,c:0}; selAnchor=null; afterChange(); renderSheets(); setDocTitle(); $("#backstage").classList.remove("show"); } };
      break;
    case "Open":
      c.innerHTML = `<h2 style="font-size:18px;margin-bottom:12px">Open</h2>
        <button class="btn primary" id="bs-open">Browse…</button>
        <p style="margin-top:14px;font-size:12.5px;color:var(--muted)">Or drag an .xlsx / .csv file onto the window.</p>`;
      $("#bs-open").onclick = ()=>{ $("#backstage").classList.remove("show"); openPicker(); };
      break;
    case "Save":
      c.innerHTML = `<h2 style="font-size:18px;margin-bottom:12px">Save</h2>
        <p style="font-size:13px;color:var(--muted);margin-bottom:12px">Your edits are already saved in this browser. Use Save As to write an .xlsx file.</p>
        <button class="btn primary" id="bs-save">Download .xlsx</button>`;
      $("#bs-save").onclick = ()=>downloadXlsx();
      break;
    case "Save As":
      c.innerHTML = `<h2 style="font-size:18px;margin-bottom:12px">Save As</h2>
        <div class="row" style="display:flex;gap:8px;align-items:center;margin-bottom:12px"><label style="min-width:60px;font-size:12.5px;color:var(--muted)">Name</label>
        <input type="text" id="bs-name" value="${name}" style="flex:1;padding:4px 8px;border:1px solid var(--border);border-radius:2px;background:var(--cell-bg);color:var(--fg)"></div>
        <button class="btn primary" id="bs-xlsx">Save as .xlsx</button>
        <button class="btn" id="bs-csv" style="margin-left:8px">Save active sheet as .csv</button>`;
      $("#bs-xlsx").onclick = ()=>{ wb.name = $("#bs-name").value.trim()||"Book1"; setDocTitle(); downloadXlsx(); };
      $("#bs-csv").onclick = ()=>{ wb.name = $("#bs-name").value.trim()||"Book1"; setDocTitle(); exportCsv(); };
      break;
    case "Print":
      c.innerHTML = `<h2 style="font-size:18px;margin-bottom:12px">Print</h2>
        <p style="font-size:13px;color:var(--muted);margin-bottom:12px">Prints the used range of the active sheet — gridlines and app chrome are excluded.</p>
        <button class="btn primary" id="bs-print">Print / Save as PDF</button>`;
      $("#bs-print").onclick = ()=>{ $("#backstage").classList.remove("show"); setTimeout(()=>printSheet(),120); };
      break;
    case "Export":
      c.innerHTML = `<h2 style="font-size:18px;margin-bottom:12px">Export</h2>
        <button class="btn primary" id="bs-ex1">Export as .xlsx</button>
        <button class="btn" id="bs-ex2" style="margin-left:8px">Export active sheet as .csv</button>
        <p style="margin-top:14px;font-size:12.5px;color:var(--muted)">Cell styling (fill, borders, fonts) is not written to the .xlsx by the bundled writer — values, formulas, merges and number formats are.</p>`;
      $("#bs-ex1").onclick = ()=>downloadXlsx();
      $("#bs-ex2").onclick = ()=>exportCsv();
      break;
    case "Account":
      c.innerHTML = `<h2 style="font-size:18px;margin-bottom:12px">Account</h2>
        <p style="font-size:13px;color:var(--muted)">This workbook runs entirely on this device. No account is connected and no data leaves the browser.</p>`;
      break;
    case "Options":
      c.innerHTML = `<h2 style="font-size:18px;margin-bottom:12px">Options</h2>
        <div class="row" style="display:flex;align-items:center;gap:8px;margin:8px 0"><label style="min-width:150px;font-size:12.5px;color:var(--muted)">Theme</label>
        <select id="bs-theme" style="padding:3px 6px;border:1px solid var(--border);border-radius:2px;background:var(--cell-bg);color:var(--fg)">
          <option value="light">Office (Light)</option><option value="dark">Dark Gray</option></select></div>
        <div class="row" style="display:flex;align-items:center;gap:8px;margin:8px 0"><label style="min-width:150px;font-size:12.5px;color:var(--muted)">Show formula bar</label>
        <input type="checkbox" id="bs-fxbar" ${FXBAR_ON?"checked":""}></div>
        <div class="row" style="display:flex;align-items:center;gap:8px;margin:8px 0"><label style="min-width:150px;font-size:12.5px;color:var(--muted)">Show headings</label>
        <input type="checkbox" id="bs-head" ${HEADINGS_ON?"checked":""}></div>`;
      $("#bs-theme").value = document.documentElement.getAttribute("data-theme")||"light";
      $("#bs-theme").onchange = e=>setTheme(e.target.value);
      $("#bs-fxbar").onchange = ()=>runCmd("formulabar");
      $("#bs-head").onchange = ()=>runCmd("headings");
      break;
    case "Close":
      $("#backstage").classList.remove("show");
      break;
  }
}
document.addEventListener("click", e=>{
  const go = e.target.closest("[data-bsgo]");
  if(go){ bsSel = go.dataset.bsgo; renderBackstage(); return; }
  const b = e.target.closest("[data-bs]");
  if(b){ bsSel=b.dataset.bs; renderBackstage(); }
});
$("#bs-close").onclick = ()=>$("#backstage").classList.remove("show");

/* ---------- keyboard ---------- */
document.addEventListener("keydown", e=>{
  const inFx = e.target.id==="fxinput";
  const inCell = e.target.id==="cellinput";
  const inDialog = (e.target && e.target.closest) ? (e.target.closest("dialog") || e.target.closest("#backstage")) : null;
  const mod = e.ctrlKey || e.metaKey;
  if(inDialog) return;

  if(inCell){
    if(e.key==="Enter"){ e.preventDefault(); commitEdit(e.shiftKey?"up":"down"); }
    else if(e.key==="Tab"){ e.preventDefault(); commitEdit("right"); }
    else if(e.key==="Escape"){ e.preventDefault(); cancelEdit(); }
    return;
  }
  if(inFx){
    if(e.key==="Enter"){ e.preventDefault(); commitFx(); }
    else if(e.key==="Escape"){ e.preventDefault(); $("#fxinput").value = rawOf(cur.r,cur.c); $("#gridwrap").focus?.(); }
    return;
  }
  if(mod){
    const k = e.key.toLowerCase();
    if(k==="z"){ e.preventDefault(); e.shiftKey? redo() : undo(); return; }
    if(k==="y"){ e.preventDefault(); redo(); return; }
    if(k==="c"){ e.preventDefault(); doCopy(false); return; }
    if(k==="x"){ e.preventDefault(); doCopy(true); return; }
    if(k==="v"){ e.preventDefault(); doPaste(); return; }
    if(k==="b"){ e.preventDefault(); runCmd("bold"); return; }
    if(k==="i"){ e.preventDefault(); runCmd("italic"); return; }
    if(k==="u"){ e.preventDefault(); runCmd("underline"); return; }
    if(k==="s"){ e.preventDefault(); downloadXlsx(); return; }
    if(k==="o"){ e.preventDefault(); openPicker(); return; }
    if(k==="p"){ e.preventDefault(); printSheet(); return; }
    if(k==="f"){ e.preventDefault(); findDialog(); return; }
    if(k==="h"){ e.preventDefault(); findDialog(); return; }
    if(k==="g"){ e.preventDefault(); gotoDialog(); return; }
    if(k==="1"){ e.preventDefault(); fmtDialog(); return; }
    if(k==="a"){ e.preventDefault(); selAnchor={r:0,c:0}; cur={...cur, r:NROWS-1, c:NCOLS-1}; positionSel(); updateStatus(); return; }
    if(k==="home"){ e.preventDefault(); cur={...cur, r:0, c:0}; selAnchor=null; positionSel(); scrollCellIntoView(0,0); return; }
    if(e.shiftKey && e.key==="%"){ e.preventDefault(); runCmd("percent"); return; }
    if(e.shiftKey && e.key==="$"){ e.preventDefault(); runCmd("currency"); return; }
    return;
  }
  switch(e.key){
    case "ArrowDown": e.preventDefault(); moveCur(1,0,e.shiftKey); break;
    case "ArrowUp": e.preventDefault(); moveCur(-1,0,e.shiftKey); break;
    case "ArrowLeft": e.preventDefault(); moveCur(0,-1,e.shiftKey); break;
    case "ArrowRight": e.preventDefault(); moveCur(0,1,e.shiftKey); break;
    case "Tab": e.preventDefault(); moveCur(0,e.shiftKey?-1:1,false); break;
    case "Enter": e.preventDefault(); moveCur(e.shiftKey?-1:1,0,false); break;
    case "F2": e.preventDefault(); startEdit(); break;
    case "F4": e.preventDefault(); runCmd("recalc"); break;
    case "F9": e.preventDefault(); runCmd("recalc"); break;
    case "Delete": e.preventDefault(); clearSelContent(); break;
    case "Backspace": e.preventDefault(); startEdit(""); break;
    case "Escape": selAnchor=null; positionSel(); break;
    case "Home": e.preventDefault(); cur = {...cur, c:0}; selAnchor=null; positionSel(); break;
    case "End": e.preventDefault(); cur = {...cur, c:NCOLS-1}; selAnchor=null; positionSel(); break;
    case "PageDown": e.preventDefault(); moveCur(25,0,false); break;
    case "PageUp": e.preventDefault(); moveCur(-25,0,false); break;
    default:
      if(e.key.length===1 && !e.altKey && !e.metaKey){ e.preventDefault(); startEdit(e.key); }
  }
});
/* fx bar wiring */
$("#fxinput").addEventListener("focus", ()=>{ if(!editing){ editing = true; editFromFx = true; } });
$("#fxinput").addEventListener("keydown", e=>{
  if(e.key==="Enter"){ e.preventDefault(); commitFx(); }
});
function commitFx(){
  const val = $("#fxinput").value;
  editing = false;
  setCellInput(cur.r, cur.c, val);
  renderGrid(); syncToolbar();
}
$("#fxok").onclick = ()=>commitFx();
$("#fxcancel").onclick = ()=>{ editing=false; $("#fxinput").value = rawOf(cur.r,cur.c); };
$("#fxfx").onclick = ()=>fxDialog();
$("#namesel").addEventListener("change", ()=>{
  const p = parseRef($("#namesel").value);
  if(p){ cur = {...cur, ...clampRC(p.r,p.c)}; selAnchor=null; positionSel(); scrollCellIntoView(p.r,p.c); }
});
/* zoom slider + view buttons */
$("#zoomslider").addEventListener("input", e=>setZoom(parseInt(e.target.value,10)));
$$("[data-view]").forEach(b=>b.addEventListener("click", ()=>{
  $$("[data-view]").forEach(x=>x.classList.toggle("on", x===b));
  if(b.dataset.view==="normal") runCmd("viewnormal");
  else if(b.dataset.view==="break") runCmd("viewpagebreak");
  else runCmd("zoomdialog");
}));

/* ---------- init ---------- */
function boot(){
  try{
    const th = localStorage.getItem("excel-web-theme");
    if(th) document.documentElement.setAttribute("data-theme", th);
  }catch(e){}

  const many = loadLocalAll();
  if(many){
    WORKBOOKS = many.books.map(b=>({
      id: b.id || newWBID(),
      name: b.name || "Book1",
      sheets: (b.sheets && b.sheets.length) ? b.sheets : [blankSheet("Sheet1")],
      active: b.active || 0,
      cur: b.cur || {sheet:0,r:0,c:0},
      dirty: false,
      savedTs: b.savedTs || 0,
      path: b.path || null,
    }));
    NEXT_ID = Math.max(many.next || 1, ...WORKBOOKS.map(w=>w.id+1));
    try{ ACTIVE = Math.min(parseInt(localStorage.getItem(LS_ACTIVE)||"0",10), WORKBOOKS.length-1); }catch(e){ ACTIVE = 0; }
    if(ACTIVE<0) ACTIVE = 0;
  }else{
    const legacy = loadLocal();
    const w = legacy && legacy.wb ? legacy.wb : newWorkbook();
    w.name = (legacy && legacy.wb && legacy.wb.name) || "Book1";
    WORKBOOKS = [{
      id: newWBID(),
      name: w.name,
      sheets: (w.sheets && w.sheets.length) ? w.sheets : [blankSheet("Sheet1")],
      active: 0,
      cur: (legacy && legacy.cur) || {sheet:0,r:0,c:0},
      dirty: false, savedTs: 0, path: null,
    }];
    ACTIVE = 0;
  }
  WORKBOOKS.forEach(w=>{ w.sheets.forEach(s=>normalizeSheet(s)); if(!w.sheets.length) w.sheets=[blankSheet("Sheet1")]; });
  syncAliases();
  selAnchor = null;

  renderTabs(); showTab("Home");
  recalcAll(); renderAll();
  initSplitDivider();
  initWBTouchScroll();
  applyPageSetup(); applyChrome();
  const a = $("#autosave"); if(a) a.textContent = many ? "Recovered" : "Saved to this device";
  const tog = $("#astoggle");
  if(tog) tog.onclick = ()=>{
    AUTOSAVE_ON = !AUTOSAVE_ON;
    tog.setAttribute("aria-pressed", AUTOSAVE_ON ? "true" : "false");
    const st = $("#autosave");
    if(AUTOSAVE_ON){ if(st) st.textContent = "Saved to this device"; saveLocal(); }
    else if(st) st.textContent = "Unsaved changes";
    flashStatus(AUTOSAVE_ON ? "AutoSave on" : "AutoSave off");
  };
  $("#st-sel").textContent = "Ready";
}
boot();

/* Service worker only over a secure origin — file:// and http:// would break it. */
if("serviceWorker" in navigator && location.protocol === "https:"){
  window.addEventListener("load", ()=>{
    navigator.serviceWorker.register("sw.js").catch(()=>{});
  });
}

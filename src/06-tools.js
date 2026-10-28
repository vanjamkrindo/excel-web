/* ============ 06 · Dialogs & Data Tools ============ */
function dlg(id, title, bodyHtml, onOk, okLabel){
  let d = document.getElementById(id);
  if(d) d.remove();
  d = document.createElement("dialog");
  d.id = id;
  d.innerHTML = `<h3>${title}</h3><div class="dbody">${bodyHtml}</div>
    <div class="actions"><button class="btn" data-cancel>Cancel</button><button class="btn primary" data-ok>${okLabel||"OK"}</button></div>`;
  document.body.appendChild(d);
  d.addEventListener("click", e=>{
    if(e.target.hasAttribute("data-cancel") || e.target===d) d.close();
    if(e.target.hasAttribute("data-ok")){ if(onOk(d)!==false) d.close(); }
  });
  d.showModal();
  const first = d.querySelector("input,select");
  if(first) setTimeout(()=>first.focus(),30);
  return d;
}
function flashStatus(msg){
  const el = $("#autosave"); if(!el) return;
  el.textContent = msg;
  clearTimeout(flashStatus._t);
  flashStatus._t = setTimeout(()=>{ el.textContent = dirty?"Unsaved changes":"Saved to this browser"; }, 1800);
}
function promptNum(label, val, cb){
  dlg("dlg-prompt", label, `<div class="row"><label>Value</label><input type="number" id="pv" value="${val}" min="1"></div>`,
    d=>{ const v = parseInt(d.querySelector("#pv").value,10); if(!v||v<1) return false; cb(v); });
}

/* --- borders --- */
function borderDialog(){
  dlg("dlg-border","Format Cells — Border",`<div class="row"><label>Presets</label><select id="bd">
    <option value="all">All borders</option><option value="outline">Outside borders</option>
    <option value="top">Top border</option><option value="bottom">Bottom border</option>
    <option value="left">Left border</option><option value="right">Right border</option><option value="none">No border</option>
  </select></div>`, d=>{
    const mode = d.querySelector("#bd").value;
    pushUndo("border");
    const s = selRange();
    eachSelCell((r,c)=>{
      const cl=getCellW(r,c); const st=cl.s||{};
      if(mode==="none"){ delete st.border; cl.s=st; return; }
      let b = st.border||0;
      const set=(bit,on)=>{ b = on ? (b|bit) : (b & ~bit); };
      if(mode==="all") b=15;
      if(mode==="outline"){
        set(1, r===s.r1); set(4, r===s.r2); set(8, c===s.c1); set(2, c===s.c2);
      }
      if(mode==="top") set(1,true);
      if(mode==="bottom") set(4,true);
      if(mode==="left") set(8,true);
      if(mode==="right") set(2,true);
      st.border=b; cl.s=st;
    });
    renderGrid(); markDirty();
  });
}
/* --- decimals --- */
function decDialog(){
  dlg("dlg-dec","Number Format — Decimals",`<div class="row"><label>Decimal places</label><input type="number" id="dn" value="2" min="0" max="10"></div>`,
    d=>{
      const n = parseInt(d.querySelector("#dn").value,10)||0;
      pushUndo("decimals");
      eachSelCell((r,c)=>{ const cl=getCellW(r,c); cl.numfmt="Dec"+n; });
      renderGrid(); markDirty();
    });
}
/* --- insert/delete cells --- */
function insDelDialog(mode){
  dlg("dlg-insdel", mode==="insert"?"Insert Cells":"Delete Cells",
  `<div class="row"><label>Direction</label><select id="dir">
    <option value="row">${mode==="insert"?"Shift cells down":"Shift cells up"}</option>
    <option value="col">${mode==="insert"?"Shift cells right":"Shift cells left"}</option>
    <option value="rowall">Entire row</option>
    <option value="colall">Entire column</option>
  </select></div>`, d=>{
    const dir = d.querySelector("#dir").value;
    if(mode==="insert"){
      if(dir==="row"||dir==="rowall") insertRows(); else insertCols();
    } else {
      if(dir==="row"||dir==="rowall") deleteRows(); else deleteCols();
    }
  });
}
/* --- format cells --- */
function fmtDialog(){
  const st = getStyle(cur.r,cur.c);
  dlg("dlg-fmt","Format Cells",`
    <div class="row"><label>Category</label><select id="cat">
      ${Object.entries(NUMFMTS).map(([k,v])=>`<option value="${k}"${numFmtOf(cur.r,cur.c)===k?" selected":""}>${v.label}</option>`).join("")}
    </select></div>
    <div class="row"><label>Font style</label>
      <label style="display:flex;gap:4px;align-items:center"><input type="checkbox" id="fb" ${st.bold?"checked":""}>Bold</label>
      <label style="display:flex;gap:4px;align-items:center"><input type="checkbox" id="fi" ${st.italic?"checked":""}>Italic</label>
      <label style="display:flex;gap:4px;align-items:center"><input type="checkbox" id="fu" ${st.underline?"checked":""}>Underline</label>
    </div>
    <div class="row"><label>Font</label><select id="ff">${FONT_LIST.map(f=>`<option${(st.font||"Calibri")===f?" selected":""}>${f}</option>`).join("")}</select>
      <select id="fs" style="max-width:70px">${SIZE_LIST.map(s=>`<option${String(st.size||11)===s?" selected":""}>${s}</option>`).join("")}</select></div>
    <div class="row"><label>Fill color</label><input type="color" id="fbg" value="${st.bg||"#ffffff"}"></div>
    <div class="row"><label>Font color</label><input type="color" id="ffg" value="${st.color||"#000000"}"></div>
    <div class="row"><label>Alignment</label><select id="fal">
      ${["left","center","right"].map(a=>`<option${st.align===a?" selected":""}>${a}</option>`).join("")}</select>
      <select id="fva">${["top","middle","bottom"].map(a=>`<option${st.valign===a?" selected":""}>${a}</option>`).join("")}</select></div>
    <div class="row"><label>Border</label><select id="fbd">
      <option value="">None</option><option value="15">All borders</option><option value="16">Outside borders</option>
      <option value="1">Top</option><option value="4">Bottom</option><option value="8">Left</option><option value="2">Right</option></select></div>`, d=>{
    const cat = d.querySelector("#cat").value;
    const bold = d.querySelector("#fb").checked, ital = d.querySelector("#fi").checked, und = d.querySelector("#fu").checked;
    const bg = d.querySelector("#fbg").value, fg = d.querySelector("#ffg").value;
    const al = d.querySelector("#fal").value, va = d.querySelector("#fva").value;
    const font = d.querySelector("#ff").value, size = parseInt(d.querySelector("#fs").value,10);
    const bd = d.querySelector("#fbd").value;
    pushUndo("format cells");
    eachSelCell((r,c)=>{
      const cl=getCellW(r,c);
      cl.numfmt=cat;
      const s = {...(cl.s||{}), bold, italic:ital, underline:und, bg, color:fg, align:al, valign:va, font, size};
      if(bd==="") delete s.border; else s.border = parseInt(bd,10);
      cl.s = s;
    });
    renderGrid(); markDirty();
  }, "OK");
}
/* --- insert function --- */
const FX_HELP = {
  SUM:"SUM(number1, [number2], …)", AVERAGE:"AVERAGE(number1, …)", IF:"IF(logical_test, value_if_true, value_if_false)",
  VLOOKUP:"VLOOKUP(lookup_value, table_array, col_index_num, 0)", SUMIF:"SUMIF(range, criteria, sum_range)",
  COUNTIF:"COUNTIF(range, criteria)", CONCAT:"CONCAT(text1, …)", ROUND:"ROUND(number, num_digits)",
  TODAY:"TODAY()", NOW:"NOW()", LEN:"LEN(text)", UPPER:"UPPER(text)", MAX:"MAX(…)", MIN:"MIN(…)",
  MEDIAN:"MEDIAN(…)", STDEV:"STDEV(…)", PMT:"PMT(rate, nper, pv)",
};
function fxDialog(){
  const names = Object.keys(FUNCS).filter(k=>typeof FUNCS[k]==="function").sort();
  dlg("dlg-fx","Insert Function",`
    <div class="row"><label>Search</label><input type="text" id="fq" placeholder="e.g. SUM"></div>
    <div class="row"><label>Select a function</label><select id="fsel" size="9" style="flex:1">
      ${names.map(n=>`<option value="${n}">${n}</option>`).join("")}</select></div>
    <div id="fhelp" style="font-size:12px;color:var(--muted);min-height:30px;padding:6px;border:1px solid var(--border-soft)"></div>`, d=>{
    const name = d.querySelector("#fsel").value;
    if(!name) return false;
    insertAtCursor(name+"(");
  }, "OK");
  const d = $("#dlg-fx");
  const sel = d.querySelector("#fsel");
  const help = d.querySelector("#fhelp");
  const upd = ()=>{ help.textContent = FX_HELP[sel.value] || (sel.value+"(…)"); };
  sel.addEventListener("change", upd); upd();
  d.querySelector("#fq").addEventListener("input", e=>{
    const q = e.target.value.toUpperCase();
    [...sel.options].forEach(o=>{ o.style.display = o.value.includes(q)?"":"none"; });
  });
}
function insertAtCursor(text){
  const inp = $("#fxinput");
  inp.focus();
  const pos = inp.selectionStart ?? inp.value.length;
  inp.value = inp.value.slice(0,pos)+text+inp.value.slice(pos);
  inp.selectionStart = inp.selectionEnd = pos+text.length;
  editing = true;
  $("#cellinput").style.display="none";
}
function autoSum(fn){
  const s = selRange();
  const r0 = s.r1, c = s.c1;
  let top = r0-1, n=0;
  while(top>=0 && typeof evalCell(top,c)==="number"){ top--; n++; }
  const f = `=${fn||"SUM"}(${cellRef(top+1,c)}:${cellRef(r0-1,c)})`;
  setCellInput(r0, c, f); renderGrid();
}
/* --- symbols --- */
function symDialog(){
  const syms = ["©","®","™","€","£","¥","₩","°","±","×","÷","≤","≥","≠","≈","√","∞","µ","Ω","π","→","←","↑","↓","•","§","¶","†","‡","‰","₽","₿","★","☆","✓","✗","♥","♦","♣","♠"];
  dlg("dlg-sym","Symbol",`<div style="display:grid;grid-template-columns:repeat(10,34px);gap:4px">
    ${syms.map(s=>`<button class="btn" data-sym="${s}" style="padding:6px 0;font-size:15px">${s}</button>`).join("")}</div>`,
    ()=>false, "Close");
  $("#dlg-sym").addEventListener("click", e=>{
    const b = e.target.closest("[data-sym]");
    if(b){ setCellInput(cur.r,cur.c, (rawOf(cur.r,cur.c)||"")+b.dataset.sym); renderGrid(); }
  });
}

/* --- sort --- */
function sortSel(asc){
  const s = selRange();
  if(s.r1===s.r2 && s.c1===s.c2){ flashStatus("Select more than one row"); return; }
  pushUndo("sort");
  const sh = sheet();
  const rows = [];
  for(let r=s.r1;r<=s.r2;r++){
    const row = [];
    for(let c=0;c<NCOLS;c++) row.push(sh.cells[ck(r,c)] ? JSON.parse(JSON.stringify(sh.cells[ck(r,c)])) : null);
    rows.push(row);
  }
  const keyOf = row => {
    const cl = row[s.c1];
    if(!cl) return "";
    if(cl.f) return "";
    return cl.v ?? "";
  };
  rows.sort((a,b)=>{
    const x=keyOf(a), y=keyOf(b);
    if(typeof x==="number"&&typeof y==="number") return asc? x-y : y-x;
    const sx=String(x).toLowerCase(), sy=String(y).toLowerCase();
    return asc ? sx.localeCompare(sy,"en") : sy.localeCompare(sx,"en");
  });
  rows.forEach((row,i)=>{ const r=s.r1+i; for(let c=0;c<NCOLS;c++){ if(row[c]) sh.cells[ck(r,c)]=row[c]; else delete sh.cells[ck(r,c)]; }});
  recalcAll(); renderGrid(); markDirty();
}
/* --- filter (AutoFilter on the active column) --- */
function toggleFilter(){
  const s = selRange();
  const sh = sheet();
  if(sh.filter){ sh.filter=null; renderGrid(); flashStatus("Filter cleared"); return; }
  const vals = new Set();
  for(let r=s.r1+1;r<=s.r2;r++){ const v = evalCell(r,s.c1); vals.add(String(v)); }
  const list = [...vals].slice(0,50);
  dlg("dlg-filter","Filter — Column "+colName(s.c1),
    `<div style="max-height:240px;overflow:auto">${list.map((v,i)=>
      `<label class="chkline" style="display:flex"><input type="checkbox" class="fv" value="${esc(v)}" checked> <span>${esc(v)||"(Blanks)"}</span></label>`).join("")}</div>`,
    d=>{
      const keep = new Set([...d.querySelectorAll(".fv")].filter(c=>c.checked).map(c=>c.value));
      const hide = new Set();
      for(let r=s.r1+1;r<=s.r2;r++){ if(!keep.has(String(evalCell(r,s.c1)))) hide.add(r); }
      sheet().filter = {col:s.c1, rows:[...hide]};
      applyFilter(); flashStatus(`Filter applied: ${hide.size} row(s) hidden`);
    }, "OK");
}
function applyFilter(){
  const f = sheet().filter;
  if(!f) return;
  const hide = new Set(f.rows);
  $$("#gridtable tbody tr").forEach(tr=>{
    const th = tr.querySelector("th"); const r = +th.dataset.row;
    tr.style.display = hide.has(r) ? "none" : "";
  });
}

/* --- find & replace --- */
function findDialog(){
  dlg("dlg-find","Find and Replace",`
    <div class="row"><label>Find what</label><input type="text" id="fq"></div>
    <div class="row"><label>Replace with</label><input type="text" id="fr"></div>
    <div class="row"><label>Options</label><label class="chkline" style="display:flex;gap:4px"><input type="checkbox" id="fc"> Match case</label></div>
    <div id="findres"></div>
    <div class="actions" style="justify-content:flex-start;margin-top:8px;padding:0;border:none;background:none">
      <button class="btn" id="fnext">Find Next</button>
      <button class="btn" id="frep">Replace</button>
      <button class="btn" id="fall">Replace All</button>
    </div>`, ()=>false, "Close");
  const d = $("#dlg-find");
  let lastHit = null;
  const matches = ()=>{
    const q = d.querySelector("#fq").value;
    const cs = d.querySelector("#fc").checked;
    if(!q) return [];
    const out=[]; const sh=sheet();
    for(const k in sh.cells){
      const [r,c]=k.split(",").map(Number);
      const txt = String(sh.cells[k].f ?? sh.cells[k].v ?? "");
      const hay = cs?txt:txt.toLowerCase(), needle = cs?q:q.toLowerCase();
      if(hay.includes(needle)) out.push({r,c});
    }
    return out.sort((a,b)=>a.r-b.r||a.c-b.c);
  };
  const goTo = ()=>{
    const m = matches(); if(!m.length){ d.querySelector("#findres").textContent="Excel cannot find the data you're searching for."; return null; }
    const idx = m.findIndex(x=> lastHit && x.r===lastHit.r && x.c===lastHit.c);
    const nx = m[(idx+1)%m.length]; lastHit = nx;
    cur = {...cur, r:nx.r, c:nx.c}; selAnchor=null; positionSel(); scrollCellIntoView(nx.r,nx.c);
    d.querySelector("#findres").textContent = `Match ${m.findIndex(x=>x.r===nx.r&&x.c===nx.c)+1} of ${m.length} — ${cellRef(nx.r,nx.c)}`;
    return nx;
  };
  d.querySelector("#fnext").onclick = ()=>{ goTo(); };
  const replaceAt = ()=>{
    const hit = lastHit || goTo();
    if(!hit) return;
    const q = d.querySelector("#fq").value, rep = d.querySelector("#fr").value;
    if(!q) return;
    const cl = sheet().cells[ck(hit.r,hit.c)];
    if(!cl) return;
    pushUndo("replace");
    const cs = d.querySelector("#fc").checked;
    cl.v = String(cl.v ?? "").replace(new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g,"\\$&"), cs?"g":"gi"), rep);
    recalcAll(); renderGrid(); markDirty();
    d.querySelector("#findres").textContent = `Replaced in ${cellRef(hit.r,hit.c)}`;
  };
  d.querySelector("#frep").onclick = replaceAt;
  d.querySelector("#fall").onclick = ()=>{
    const q = d.querySelector("#fq").value, rep = d.querySelector("#fr").value;
    if(!q) return;
    pushUndo("replace all");
    let n=0;
    const cs = d.querySelector("#fc").checked;
    const sh=sheet();
    for(const k in sh.cells){
      const cl = sh.cells[k];
      if(cl.f) continue;
      const t = String(cl.v ?? "");
      const hay = cs?t:t.toLowerCase(), needle = cs?q:q.toLowerCase();
      if(hay.includes(needle)){
        cl.v = t.replace(new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g,"\\$&"), cs?"g":"gi"), rep);
        n++;
      }
    }
    recalcAll(); renderGrid(); markDirty();
    d.querySelector("#findres").textContent = `We made ${n} replacement(s).`;
  };
}
/* --- go to --- */
function gotoDialog(){
  dlg("dlg-goto","Go To",`<div class="row"><label>Reference</label><input type="text" id="ga" placeholder="e.g. C12 or A1:D20"></div>`, d=>{
    const v = d.querySelector("#ga").value.trim();
    const parts = v.split(":");
    const p1 = parseRef(parts[0]); if(!p1) return false;
    const p2 = parts[1] ? parseRef(parts[1]) : null;
    selAnchor = p2 ? p1 : null;
    cur = {...cur, ...(p2 || p1)};
    positionSel(); scrollCellIntoView(p1.r,p1.c);
  }, "OK");
}
/* --- remove duplicates --- */
function deleteDupes(){
  const s = selRange();
  if(s.r2-s.r1<1){ flashStatus("Select at least 2 rows"); return; }
  pushUndo("remove duplicates");
  const sh = sheet(); const seen = new Set(); const kill = [];
  for(let r=s.r1;r<=s.r2;r++){
    let sig="";
    for(let c=s.c1;c<=s.c2;c++) sig += String(evalCell(r,c))+"\u0001";
    if(seen.has(sig)) kill.push(r); else seen.add(sig);
  }
  kill.reverse().forEach(r=>{ for(let c=0;c<NCOLS;c++) delete sh.cells[ck(r,c)]; });
  recalcAll(); renderGrid(); markDirty();
  flashStatus(`${kill.length} duplicate row(s) removed`);
}
/* --- protect / comment --- */
function toggleProtect(){
  pushUndo("protect");
  sheet().protect = !sheet().protect;
  flashStatus(sheet().protect ? "Sheet protected (read-only)" : "Protection removed");
  renderGrid();
}
function commentDialog(){
  const cl = getCell(cur.r,cur.c);
  const cur_ = cl && cl.comment ? cl.comment : "";
  dlg("dlg-comment","Comment — "+cellRef(cur.r,cur.c),`<div class="row"><label>Comment</label><input type="text" id="cm" value="${esc(cur_)}"></div>`, d=>{
    pushUndo("comment");
    const cell = getCellW(cur.r,cur.c);
    const v = d.querySelector("#cm").value;
    if(v) cell.comment = v; else delete cell.comment;
    renderGrid(); markDirty();
  }, "Save");
}
function deleteComment(){
  const cl = getCell(cur.r,cur.c);
  if(!cl || !cl.comment){ flashStatus("No comment in the active cell"); return; }
  pushUndo("delete comment");
  delete cl.comment; renderGrid(); markDirty();
}
/* --- freeze panes --- */
function freezeDialog(){
  dlg("dlg-freeze","Freeze Panes",`<div class="row"><label>Mode</label><select id="fz">
    <option value="none">No frozen panes</option>
    <option value="row">Rows above the active cell (${cur.r} row(s))</option>
    <option value="col">Columns left of the active cell (${cur.c} column(s))</option>
    <option value="both">Rows and columns</option>
  </select></div>`, d=>{
    const m = d.querySelector("#fz").value;
    pushUndo("freeze");
    sheet().frozen = m==="none" ? {rows:0,cols:0} : m==="row" ? {rows:cur.r,cols:0} : m==="col" ? {rows:0,cols:cur.c} : {rows:cur.r,cols:cur.c};
    renderGrid();
  }, "OK");
}
/* --- theme --- */
function setTheme(name){
  document.documentElement.setAttribute("data-theme", name==="dark"?"dark":"light");
  try{ localStorage.setItem("excel-web-theme", name==="dark"?"dark":"light"); }catch(e){}
  drawAllCharts();
}
function toggleTheme(){
  const dark = document.documentElement.getAttribute("data-theme")==="dark";
  setTheme(dark?"light":"dark");
}
/* --- charts --- */
function insertChart(kind){
  const s = selRange();
  if(s.r1===s.r2 && s.c1===s.c2){ flashStatus("Select a data range first"); return; }
  pushUndo("chart");
  const series = [];
  const labels = [];
  for(let r=s.r1+1;r<=s.r2;r++) labels.push(String(evalCell(r,s.c1)));
  for(let c=s.c1+1;c<=s.c2;c++){
    const name = String(evalCell(s.r1,c)||colName(c));
    const data = [];
    for(let r=s.r1+1;r<=s.r2;r++){ const v=evalCell(r,c); data.push(typeof v==="number"?v:0); }
    series.push({name, data});
  }
  if(!series.length){
    const data=[]; for(let r=s.r1;r<=s.r2;r++){ const v=evalCell(r,s.c1); data.push(typeof v==="number"?v:0); labels.push(String(r-s.r1+1)); }
    series.push({name:"Series1", data});
  }
  const chart = {kind, labels, series, x:120+Math.random()*60, y:120+Math.random()*60, w:400, h:250, id:Date.now()};
  sheet().charts.push(chart);
  drawChart(chart);
  markDirty();
}
function drawChart(ch){
  const wrap = $("#gridwrap");
  let box = document.getElementById("chart"+ch.id);
  if(!box){
    box = document.createElement("div");
    box.className="chartbox"; box.id="chart"+ch.id;
    box.innerHTML = `<div class="cbar"><button data-chartclose title="Delete chart">✕</button></div><canvas width="${ch.w}" height="${ch.h}"></canvas>`;
    wrap.appendChild(box);
    box.querySelector("[data-chartclose]").onclick = ()=>{
      pushUndo("delete chart");
      sheet().charts = sheet().charts.filter(x=>x.id!==ch.id);
      box.remove(); markDirty();
    };
  }
  box.style.left = ch.x+"px"; box.style.top = ch.y+"px";
  const cv = box.querySelector("canvas");
  const ctx = cv.getContext("2d");
  const W=cv.width, H=cv.height;
  ctx.clearRect(0,0,W,H);
  const dark = document.documentElement.getAttribute("data-theme")==="dark";
  ctx.fillStyle = dark? "#1e1e1e" : "#ffffff"; ctx.fillRect(0,0,W,H);
  const pad = {l:48,r:12,t:22,b:30};
  const all = ch.series.flatMap(s=>s.data);
  const max = Math.max(1, ...all), min = Math.min(0, ...all);
  const plotW = W-pad.l-pad.r, plotH = H-pad.t-pad.b;
  const colors = ["#217346","#4472c4","#c00000","#ffc000","#7030a0","#00b0f0"];
  ctx.strokeStyle = dark? "#555":"#d4d4d4"; ctx.lineWidth=1;
  ctx.beginPath(); ctx.moveTo(pad.l,pad.t); ctx.lineTo(pad.l,H-pad.b); ctx.lineTo(W-pad.r,H-pad.b); ctx.stroke();
  ctx.fillStyle = dark? "#bbb":"#595959"; ctx.font="10px Segoe UI";
  for(let i=0;i<=4;i++){
    const y = pad.t+plotH*i/4; const val = max - (max-min)*i/4;
    ctx.fillText(fmtNum(Math.round(val*100)/100), 2, y+3);
    ctx.strokeStyle = dark? "#3a3a3a":"#f0f0f0";
    ctx.beginPath(); ctx.moveTo(pad.l,y); ctx.lineTo(W-pad.r,y); ctx.stroke();
  }
  const n = ch.labels.length||1;
  if(ch.kind==="pie"){
    const data = ch.series[0].data; const total = data.reduce((a,b)=>a+Math.abs(b),0)||1;
    let ang = -Math.PI/2; const cx=W/2, cy=H/2-6, rad=Math.min(W,H)/2-34;
    data.forEach((v,i)=>{
      const a2 = ang + Math.abs(v)/total*Math.PI*2;
      ctx.beginPath(); ctx.moveTo(cx,cy); ctx.arc(cx,cy,rad,ang,a2); ctx.closePath();
      ctx.fillStyle = colors[i%colors.length]; ctx.fill();
      ang = a2;
    });
    return;
  }
  ch.series.forEach((sr,si)=>{
    ctx.strokeStyle = colors[si%colors.length]; ctx.fillStyle = ctx.strokeStyle; ctx.lineWidth=2;
    if(ch.kind==="bar"){
      const bw = plotW/n/ch.series.length;
      sr.data.forEach((v,i)=>{
        const x = pad.l + i*(plotW/n) + si*bw;
        const y = pad.t + plotH*(max-v)/(max-min);
        ctx.fillRect(x+1, y, Math.max(2,bw-2), H-pad.b-y);
      });
    } else {
      ctx.beginPath();
      sr.data.forEach((v,i)=>{
        const x = pad.l + (n>1? plotW*i/(n-1):plotW/2);
        const y = pad.t + plotH*(max-v)/(max-min);
        i? ctx.lineTo(x,y) : ctx.moveTo(x,y);
      });
      ctx.stroke();
    }
  });
  ctx.fillStyle = dark? "#bbb":"#595959"; ctx.font="9.5px Segoe UI";
  ch.labels.forEach((l,i)=>{
    if(n>12 && i%Math.ceil(n/8)!==0) return;
    const x = pad.l + (ch.kind==="bar" ? plotW*(i+0.5)/n : plotW*(n>1? i/(n-1):0.5));
    ctx.fillText(String(l).slice(0,8), x-10, H-pad.b+12);
  });
  let lx = pad.l;
  ch.series.forEach((sr,i)=>{ ctx.fillStyle=colors[i%colors.length]; ctx.fillRect(lx, 6, 8, 8); ctx.fillStyle=dark?"#bbb":"#595959"; ctx.fillText(sr.name, lx+11, 14); lx += 14+ctx.measureText(sr.name).width; });
}
function drawAllCharts(){ (sheet().charts||[]).forEach(drawChart); }

/* --- print: used range / print area only --- */
function usedRange(){
  let maxR=-1, maxC=-1;
  for(const k in sheet().cells){
    const [r,c]=k.split(",").map(Number);
    if(r>maxR) maxR=r; if(c>maxC) maxC=c;
  }
  return {r1:0, c1:0, r2:Math.max(maxR,0), c2:Math.max(maxC,0)};
}
function printSheet(){
  const ps = sheet().pageSetup || {};
  const area = ps.printArea || usedRange();
  const {r1, c1, r2, c2} = area;
  const hideRows = [], hideCols = [];
  $$("#gridtable tbody tr").forEach(tr=>{
    const th = tr.querySelector("th");
    if(+th.dataset.row > r2 || +th.dataset.row < r1) { tr.dataset.prevDisplay = tr.style.display||""; tr.style.display="none"; hideRows.push(tr); }
  });
  $$("#gridtable thead th[data-col]").forEach(th=>{
    if(+th.dataset.col > c2 || +th.dataset.col < c1) th.style.display="none";
  });
  $$("#gridtable td").forEach(td=>{
    if(+td.dataset.c > c2 || +td.dataset.r > r2) td.style.display="none";
  });
  const restore = ()=>{
    hideRows.forEach(tr=> tr.style.display = tr.dataset.prevDisplay||"");
    $$("#gridtable thead th[data-col]").forEach(th=> th.style.display="");
    $$("#gridtable td").forEach(td=> td.style.display="");
    removeEventListener("afterprint", restore);
  };
  addEventListener("afterprint", restore);
  setTimeout(()=>{ window.print(); setTimeout(restore, 800); }, 80);
}
/* --- workbook statistics --- */
function statsDialog(){
  const sh = sheet();
  const nums = [], texts = [], formulas = [];
  for(const k in sh.cells){
    const cl = sh.cells[k];
    if(cl.f) formulas.push(k);
    const v = cl.f ? evalCache[k] : cl.v;
    if(typeof v==="number") nums.push(v);
    else if(typeof v==="string" && v) texts.push(v);
  }
  const sum = nums.reduce((a,b)=>a+b,0);
  const avg = nums.length ? sum/nums.length : 0;
  const fmt = n => n.toLocaleString("en-US",{maximumFractionDigits:4});
  dlg("dlg-stats","Workbook Statistics — "+sh.name,`
    <table style="font-size:12.5px;border-collapse:collapse">
      ${[["Filled cells", Object.keys(sh.cells).length],["Numeric cells", nums.length],["Text cells", texts.length],["Formula cells", formulas.length],
         ["Sum", fmt(sum)],["Average", fmt(avg)],["Minimum", nums.length?fmt(Math.min(...nums)):"—"],["Maximum", nums.length?fmt(Math.max(...nums)):"—"],
         ["Sheets", wb.sheets.length],["Charts", (sh.charts||[]).length]]
        .map(([k,v])=>`<tr><td style="padding:3px 16px 3px 0;color:var(--muted)">${k}</td><td style="padding:3px 0;font-weight:600">${v}</td></tr>`).join("")}
    </table>`, ()=>false, "Close");
}
function spellCheck(){
  const bad = [];
  const sh = sheet();
  const re = /\b(\w+)\s+\1\b/gi; // repeated words
  for(const k in sh.cells){
    const v = sh.cells[k].v;
    if(typeof v!=="string" || !v) continue;
    const m = v.match(re);
    if(m) { const [r,c]=k.split(",").map(Number); bad.push(cellRef(r,c)+": "+m[0]); }
  }
  dlg("dlg-spell","Spelling", bad.length
    ? `<p style="font-size:12.5px;margin-bottom:8px">Repeated words found:</p><ul style="font-size:12.5px;padding-left:18px">${bad.slice(0,20).map(b=>`<li>${esc(b)}</li>`).join("")}</ul>`
    : `<p style="font-size:12.5px">No repeated words found. This check detects repeated words only — it is not a full dictionary.</p>`,
    ()=>false, "Close");
}

/* ============ 03 · Grid render & editing ============ */
let GRID_ON = true, SHOW_FORMULAS = false, ZOOM = 100;

function displayValue(r,c){
  const cell = getCell(r,c);
  if(!cell) return "";
  if(cell.f && !SHOW_FORMULAS) return formatValue(evalCell(r,c), cell.numfmt || "General");
  if(cell.f && SHOW_FORMULAS) return cell.f;
  return formatValue(cell.v, cell.numfmt || "General");
}

function formatValue(v, fmt){
  if(v==null || v==="") return "";
  if(typeof v==="number"){
    const neg = v<0, a = Math.abs(v);
    const en = (n, min=0, max=10) => n.toLocaleString("en-US",{minimumFractionDigits:min, maximumFractionDigits:max});
    switch(fmt){
      case "Number": return (neg?"-":"")+en(a,2,2);
      case "Accounting": return (neg?"($":"$")+en(a,2,2)+(neg?")":"");
      case "Currency": return (neg?"-$":"$")+en(a,2,2);
      case "Percent": return en(v*100,2,2)+"%";
      case "Scientific": return v.toExponential(2).replace("e","E");
      case "ShortDate": return serialToDate(v).toLocaleDateString("en-US",{month:"numeric",day:"numeric",year:"numeric"});
      case "LongDate": return serialToDate(v).toLocaleDateString("en-US",{weekday:"long",month:"long",day:"numeric",year:"numeric"});
      case "Date": return serialToDate(v).toLocaleDateString("en-US",{month:"numeric",day:"numeric",year:"numeric"});
      case "Time": return serialToDate(v).toLocaleTimeString("en-US",{hour:"2-digit",minute:"2-digit",second:"2-digit"});
      case "Text": return String(v);
      default:
        if(/^Dec\d+$/.test(fmt)){ const d = +fmt.slice(3); return en(v,d,d); }
        return fmtNum(v);
    }
  }
  return String(v);
}
function idnFrac(v){
  const frac = Math.abs(v)%1;
  return ","+String(Math.round(frac*100)).padStart(2,"0").replace(/0$/,"");
}
function serialToDate(n){
  return new Date(Math.round((n-25569)*86400000));
}
function dateToSerial(d){
  return Math.round((d.getTime()/86400000)+25569 - d.getTimezoneOffset()/1440*0 + 0) ; // UTC-based below
}
function toSerial(d){
  return (Date.UTC(d.getFullYear(),d.getMonth(),d.getDate(),d.getHours(),d.getMinutes()) / 86400000) + 25569 - 0.0;
}

function colW(c){ const w = sheet().cols[c]; return w ? w : 64; }
function rowH(r){ const h = sheet().rows[r]; return h ? h : 20; }

function renderGrid(){
  const t = $("#gridtable");
  const sh = sheet();
  const showGrid = GRID_ON;
  const merged = mergeMap();
  let head = "<tr><th class='corner'></th>";
  for(let c=0;c<NCOLS;c++) head += `<th data-col="${c}" style="width:${colW(c)}px;min-width:${colW(c)}px;${sh.frozen.cols>c?"position:sticky;left:"+(44+prevW(c))+"px;z-index:6;":""}">${colName(c)}</th>`;
  head += "</tr>";
  t.querySelector("thead").innerHTML = head;
  let body = "";
  for(let r=0;r<NROWS;r++){
    body += `<tr><th data-row="${r}" style="height:${rowH(r)}px;${sh.frozen.rows>r?"position:sticky;top:0;":""}">${r+1}</th>`;
    for(let c=0;c<NCOLS;c++){
      const m = merged[ck(r,c)];
      let span = "";
      if(m){
        if(m.hide){ body += `<td data-r="${r}" data-c="${c}" class="merged-cell"></td>`; continue; }
        span = ` colspan="${m.c2-m.c1+1}" rowspan="${m.r2-m.r1+1}"`;
      }
      const st = getStyle(r,c);
      const css = [];
      if(st.bold) css.push("font-weight:700");
      if(st.italic) css.push("font-style:italic");
      if(st.underline) css.push("text-decoration:underline");
      if(st.strike) css.push("text-decoration:line-through");
      if(st.underline&&st.strike) css.push("text-decoration:underline line-through");
      if(st.color) css.push("color:"+st.color);
      if(st.bg) css.push("background:"+st.bg);
      if(st.font) css.push("font-family:"+st.font);
      if(st.size) css.push("font-size:"+st.size+"px");
      if(st.align) css.push("text-align:"+st.align);
      if(st.valign) css.push("vertical-align:"+st.valign);
      if(st.wrap) css.push("white-space:normal;word-break:break-word");
      if(st.border) css.push(borderCss(st.border));
      if(st.indent) css.push("padding-left:"+(3+st.indent*10)+"px");
      if(st.dblunder) css.push("text-decoration:underline double");
      const num = typeof evalCache[ck(r,c)] === "number" || (typeof (getCell(r,c)||{}).v === "number" && !getCell(r,c).f);
      const dv = displayValue(r,c);
      const cellObj = getCell(r,c);
      const cmt = cellObj && cellObj.comment ? `<span class="cmt" title="${esc(cellObj.comment)}"></span>` : "";
      body += `<td data-r="${r}" data-c="${c}"${span} class="${css.length?" ":""}" style="${css.join(";")}">${esc(dv)}${cmt}</td>`;
    }
    body += "</tr>";
  }
  t.querySelector("tbody").innerHTML = body;
  if(!showGrid){ t.style.borderColor="transparent"; }
  applyGridlines();
  applyCondFormats();
  applyFilter();
  applyView();
  positionSel();
  renderSheets();
  updateStatus();
}
function borderCss(b){
  const s = "1px solid #333";
  const sides = [];
  if(b&1) sides.push("border-top:"+s);
  if(b&2) sides.push("border-right:"+s);
  if(b&4) sides.push("border-bottom:"+s);
  if(b&8) sides.push("border-left:"+s);
  return sides.join(";");
}
function prevW(c){ let x=0; for(let i=0;i<c;i++) x+=colW(i); return x; }

function mergeMap(){
  const map = {};
  sheet().merges.forEach((m,mi)=>{
    for(let r=m.r1;r<=m.r2;r++) for(let c=m.c1;c<=m.c2;c++){
      map[ck(r,c)] = (r===m.r1&&c===m.c1) ? {...m} : {hide:true};
    }
  });
  return map;
}

function applyGridlines(){
  const t = $("#gridtable");
  t.querySelectorAll("td").forEach(td=>{
    if(!GRID_ON) td.style.borderColor="transparent";
    else { td.style.removeProperty("border-color"); const st=getStyle(+td.dataset.r,+td.dataset.c); if(st.border) td.style.cssText = td.style.cssText; }
  });
}

function positionSel(){
  $$("#gridtable td.sel").forEach(td=>td.classList.remove("sel"));
  $$("#gridtable td.inrange").forEach(td=>td.classList.remove("inrange"));
  const merged = mergeMap();
  const s = selRange();
  for(let r=s.r1;r<=s.r2;r++) for(let c=s.c1;c<=s.c2;c++){
    const m = merged[ck(r,c)];
    const rr = m && !m.hide ? m.r1 : r, cc = m && !m.hide ? m.c1 : c;
    const td = cellTd(rr,cc); if(!td) continue;
    if(rr===cur.r&&cc===cur.c) td.classList.add("sel");
    else td.classList.add("inrange");
  }
  $("#namesel").value = cellRef(cur.r,cur.c);
  const cell = getCell(cur.r,cur.c);
  $("#fxinput").value = cell && cell.f ? cell.f : (cell ? (cell.v ?? "") : "");
  positionFillHandle();
}
/* ---- fill handle ---- */
function positionFillHandle(){
  let h = document.getElementById("fillhandle");
  const s = selRange();
  const merged = mergeMap();
  const m = merged[ck(s.r2,s.c2)];
  const rr = m && !m.hide ? m.r2 : s.r2, cc = m && !m.hide ? m.c2 : s.c2;
  const td = cellTd(rr,cc);
  if(!td){ if(h) h.style.display="none"; return; }
  if(!h){
    h = document.createElement("div");
    h.id="fillhandle"; h.className="fillhd";
    $("#gridwrap").appendChild(h);
    h.addEventListener("mousedown", startFillDrag);
  }
  h.style.display="block";
  h.style.left = (td.offsetLeft + td.offsetWidth - 4)+"px";
  h.style.top  = (td.offsetTop + td.offsetHeight - 4)+"px";
}
let fillDrag = null;
function startFillDrag(e){
  e.stopPropagation(); e.preventDefault();
  const s = selRange();
  fillDrag = {base:{...s}, to:{r:s.r2,c:s.c2}};
  const move = ev=>{
    const el = document.elementFromPoint(ev.clientX, ev.clientY);
    const td = el && el.closest && el.closest("td");
    if(td && td.dataset.r!==undefined){
      fillDrag.to = {r:+td.dataset.r, c:+td.dataset.c};
      previewFill(fillDrag);
    }
  };
  const up = ()=>{
    document.removeEventListener("mousemove", move);
    document.removeEventListener("mouseup", up);
    clearFillPreview();
    applyFill(fillDrag);
    fillDrag = null;
  };
  document.addEventListener("mousemove", move);
  document.addEventListener("mouseup", up);
}
function fillTargets(fd){
  const b = fd.base, t = fd.to;
  const out = [];
  if(t.r>b.r2){ for(let r=b.r2+1;r<=t.r;r++) for(let c=b.c1;c<=b.c2;c++) out.push({r,c,sr:b.r1+(r-(b.r2+1))%Math.max(1,b.r2-b.r1+1), sc:c}); }
  if(t.r<b.r1){ for(let r=t.r;r<b.r1;r++) for(let c=b.c1;c<=b.c2;c++) out.push({r,c,sr:b.r1+(r-b.r1)%Math.max(1,b.r2-b.r1+1), sc:c}); }
  if(t.c>b.c2){ for(let c=b.c2+1;c<=t.c;c++) for(let r=b.r1;r<=b.r2;r++) out.push({r,c,sr:r, sc:b.c1+(c-(b.c2+1))%Math.max(1,b.c2-b.c1+1)}); }
  if(t.c<b.c1){ for(let c=t.c;c<b.c1;c++) for(let r=b.r1;r<=b.r2;r++) out.push({r,c,sr:r, sc:b.c1+(c-b.c1)%Math.max(1,b.c2-b.c1+1)}); }
  return out;
}
function clearFillPreview(){
  $$("#gridtable td.fillprev").forEach(td=>{ td.classList.remove("fillprev"); td.style.boxShadow=""; });
}
function previewFill(fd){
  clearFillPreview();
  fillTargets(fd).forEach(t=>{
    const td = cellTd(t.r,t.c);
    if(td){ td.classList.add("fillprev"); td.style.boxShadow="inset 0 0 0 1px var(--sel)"; }
  });
}
function applyFill(fd){
  if(!fd) return;
  const targets = fillTargets(fd);
  if(!targets.length) return;
  pushUndo("isi cepat");
  const sh = sheet();
  // deteksi deret angka (>=2 sel) untuk ekstrapolasi
  const b = fd.base;
  const vertical = targets[0].r !== undefined && (targets[0].r > b.r2 || targets[0].r < b.r1);
  const seriesLen = vertical ? (b.r2-b.r1+1) : (b.c2-b.c1+1);
  let step = null;
  if(seriesLen>=2){
    const a1 = vertical ? evalCell(b.r2-1,b.c1) : evalCell(b.r1,b.c2-1);
    const a2 = vertical ? evalCell(b.r2,b.c1)   : evalCell(b.r1,b.c2);
    if(typeof a1==="number" && typeof a2==="number") step = a2-a1;
  }
  targets.forEach(t=>{
    const src = sh.cells[ck(t.sr,t.sc)];
    if(!src){ delete sh.cells[ck(t.r,t.c)]; return; }
    const copy = JSON.parse(JSON.stringify(src));
    if(copy.f){
      const dR = t.r - t.sr, dC = t.c - t.sc;
      copy.f = shiftFormula(copy.f, dR, dC);
    } else if(step!==null && typeof copy.v==="number"){
      const idx = vertical ? (t.r-b.r2) : (t.c-b.c2);
      const base = vertical ? evalCell(b.r2,b.c1) : evalCell(b.r1,b.c2);
      copy.v = base + step*idx;
    }
    sh.cells[ck(t.r,t.c)] = copy;
  });
  cur = {...cur, r:Math.max(fd.to.r, b.r2), c:Math.max(fd.to.c, b.c2)};
  recalcAll(); renderGrid(); markDirty();
}
function shiftFormula(f, dR, dC){
  return f.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, (m,d1,col,d2,row)=>{
    const c = colIdx(col), r = parseInt(row,10)-1;
    const nc = d1 ? c : c + dC, nr = d2 ? r : r + dR;
    if(nc<0||nr<0) return "#REF!";
    return (d1?"$":"")+colName(nc)+(d2?"$":"")+(nr+1);
  });
}
function cellTd(r,c){
  return document.querySelector(`#gridtable td[data-r="${r}"][data-c="${c}"]:not(.merged-cell)`);
}

/* selection state for toolbar */
function syncToolbar(){
  const st = getStyle(cur.r,cur.c);
  $$('[data-cmd="bold"].rbtn').forEach(b=>b.classList.toggle("on",!!st.bold));
  $$('[data-cmd="italic"].rbtn').forEach(b=>b.classList.toggle("on",!!st.italic));
  $$('[data-cmd="underline"].rbtn').forEach(b=>b.classList.toggle("on",!!st.underline));
  $$('[data-cmd="strike"].rbtn').forEach(b=>b.classList.toggle("on",!!st.strike));
  $$('[data-cmd="wrap"].rbtn').forEach(b=>b.classList.toggle("on",!!st.wrap));
  $$('[data-cmd="merge"].rbtn').forEach(b=>b.classList.toggle("on", isMergedCell(cur.r,cur.c)));
  const fn = st.font || "Calibri";
  $$('[data-cmd="fontname"]').forEach(s=>s.value=fn);
  $$('[data-cmd="fontsize"]').forEach(s=>s.value=String(st.size||11));
  $$('[data-cmd="numfmt"]').forEach(s=>s.value=numFmtOf(cur.r,cur.c));
}
function isMergedCell(r,c){
  return sheet().merges.some(m=>r>=m.r1&&r<=m.r2&&c>=m.c1&&c<=m.c2);
}

/* ---- editing ---- */
function startEdit(initial){
  if(sheet().protect) { flashStatus("This sheet is protected"); return; }
  editing = true;
  const inp = $("#cellinput");
  const cell = getCell(cur.r,cur.c);
  inp.value = initial!==undefined ? initial : (cell && cell.f ? "="+cell.f.replace(/^=/,"") : (cell && cell.v!=null ? cell.v : ""));
  inp.style.display = "block";
  const td = cellTd(cur.r,cur.c);
  if(td){
    const m = mergeMap()[ck(cur.r,cur.c)];
    const w = m && !m.hide ? td.offsetWidth : colW(cur.c);
    const h = m && !m.hide ? td.offsetHeight : rowH(cur.r);
    inp.style.left = td.offsetLeft+"px"; inp.style.top = td.offsetTop+"px";
    inp.style.width = Math.max(w,60)+"px"; inp.style.height = h+"px";
  }
  inp.focus();
  if(initial===undefined) inp.select();
}
function commitEdit(move){
  if(!editing) return;
  const inp = $("#cellinput");
  const val = inp.value;
  editing = false; inp.style.display="none";
  setCellInput(cur.r,cur.c,val);
  renderGrid();
  if(move==="down") moveCur(1,0);
  else if(move==="right") moveCur(0,1);
  else if(move==="up") moveCur(-1,0);
  $("#gridwrap").focus?.();
}
function cancelEdit(){ editing=false; $("#cellinput").style.display="none"; $("#fxinput").value = rawOf(cur.r,cur.c); $("#gridwrap").focus?.(); }
function rawOf(r,c){ const cl=getCell(r,c); return cl ? (cl.f ?? cl.v ?? "") : ""; }

function setCellInput(r,c,val){
  val = String(val ?? "");
  if(typeof validateInput==="function" && !validateInput(r,c,val)) return;
  const key = ck(r,c);
  pushUndo("edit "+cellRef(r,c));
  const sh = sheet();
  if(val===""){ delete sh.cells[key]; }
  else if(val.startsWith("=")){
    sh.cells[key] = {f:val, s:sh.cells[key]?.s};
  } else {
    let v = val;
    const t = val.trim();
    if(/^-?\d+([.,]\d+)?$/.test(t) && !/^\+?0\d/.test(t)) v = parseFloat(t.replace(",","."));
    else if(/^(3[01]|[12]\d|0?[1-9])[\/-](0?[1-9]|1[0-2])[\/-]\d{2,4}$/.test(t)){
      const p = t.split(/[\/-]/).map(Number); let y=p[2]; if(y<100) y+=2000;
      v = toSerial(new Date(y,p[1]-1,p[0]));
      sh.cells[key] = {v, numfmt:sh.cells[key]?.numfmt||"Date", s:sh.cells[key]?.s};
      recalcAll(); markDirty(); return;
    }
    sh.cells[key] = {v, s:sh.cells[key]?.s, numfmt:sh.cells[key]?.numfmt};
  }
  recalcAll(); markDirty();
}

/* ---- navigation ---- */
function moveCur(dr,dc,extend){
  const nr = cur.r+dr, nc = cur.c+dc;
  if(nr<0||nc<0||nr>=NROWS||nc>=NCOLS) return;
  if(extend){ if(!selAnchor) selAnchor={...cur}; }
  else selAnchor = null;
  cur = {...cur, r:nr, c:nc};
  positionSel(); scrollCellIntoView(nr,nc);
}
function scrollCellIntoView(r,c){
  const td = cellTd(r,c); if(!td) return;
  const wrap = $("#gridwrap");
  const b = td.getBoundingClientRect(), w = wrap.getBoundingClientRect();
  if(b.bottom>w.bottom) wrap.scrollTop += b.bottom-w.bottom+4;
  else if(b.top<w.top+22) wrap.scrollTop -= w.top+22-b.top+4;
  if(b.right>w.right) wrap.scrollLeft += b.right-w.right+4;
  else if(b.left<w.left+44) wrap.scrollLeft -= w.left+44-b.left+4;
}

/* --- ubah lebar kolom / tinggi baris dengan menyeret batas header --- */
(function(){
  const wrap = $("#gridwrap");
  let mode=null, idx=0, start=0, base=0;
  wrap.addEventListener("mousedown", e=>{
    const th = e.target.closest("th");
    if(!th) return;
    const r = th.getBoundingClientRect();
    const nearEdge = (e.clientX > r.right-5);
    if(th.dataset.col!==undefined && nearEdge){
      mode="col"; idx=+th.dataset.col; start=e.clientX; base=colW(idx);
      e.preventDefault(); e.stopPropagation();
    } else if(th.dataset.row!==undefined && e.clientY > r.bottom-5){
      mode="row"; idx=+th.dataset.row; start=e.clientY; base=rowH(idx);
      e.preventDefault(); e.stopPropagation();
    }
  }, true);
  document.addEventListener("mousemove", e=>{
    if(!mode) return;
    if(mode==="col"){
      const w = Math.max(24, base + (e.clientX-start));
      sheet().cols[idx]=w;
      const th = document.querySelector(`#gridtable th[data-col="${idx}"]`);
      if(th){ th.style.width=w+"px"; th.style.minWidth=w+"px"; }
      $$(`#gridtable td[data-c="${idx}"]`).forEach(td=>{ td.style.width=w+"px"; });
    } else {
      const h = Math.max(14, base + (e.clientY-start));
      sheet().rows[idx]=h;
      const th = document.querySelector(`#gridtable th[data-row="${idx}"]`);
      if(th) th.style.height=h+"px";
      $$(`#gridtable td[data-r="${idx}"]`).forEach(td=>{ td.style.height=h+"px"; });
    }
    e.preventDefault();
  });
  document.addEventListener("mouseup", ()=>{
    if(mode){ mode=null; renderGrid(); markDirty(); }
  });
})();

/* --- Quick Access Toolbar --- */
document.addEventListener("click", e=>{
  const b = e.target.closest("[data-qat]");
  if(!b) return;
  const a = b.dataset.qat;
  if(a==="undo") undo();
  if(a==="redo") redo();
  if(a==="save") downloadXlsx();
});

/* mouse on grid */
(function(){
  const wrap = $("#gridwrap");
  let dragging = false;
  wrap.addEventListener("mousedown", e=>{
    const th = e.target.closest("th");
    const td = e.target.closest("td");
    if(editing && td){ commitEdit(null); }
    if(th && th.dataset.col!==undefined){ selectCol(+th.dataset.col, e.shiftKey); return; }
    if(th && th.dataset.row!==undefined){ selectRow(+th.dataset.row, e.shiftKey); return; }
    if(!td) return;
    dragging = true;
    const r=+td.dataset.r, c=+td.dataset.c;
    if(typeof painter!=="undefined" && painter){
      cur = {...cur, r, c}; selAnchor=null;
      applyPainter(r,c); dragging=false; return;
    }
    if(e.shiftKey){ selAnchor=selAnchor||{...cur}; }
    else selAnchor=null;
    cur={...cur, r, c};
    positionSel();
    e.preventDefault();
  });
  wrap.addEventListener("mousemove", e=>{
    if(!dragging) return;
    const td = e.target.closest("td"); if(!td) return;
    cur = {...cur, r:+td.dataset.r, c:+td.dataset.c};
    positionSel();
  });
  addEventListener("mouseup", ()=>{ dragging=false; });
  wrap.addEventListener("dblclick", e=>{
    const td = e.target.closest("td"); if(!td) return;
    cur={...cur, r:+td.dataset.r, c:+td.dataset.c}; selAnchor=null;
    startEdit();
  });
  wrap.addEventListener("contextmenu", e=>{
    const td = e.target.closest("td");
    if(td){ e.preventDefault(); showCellMenu(e.clientX,e.clientY); }
  });
})();

function selectCol(c, add){
  selAnchor = {r:0,c};
  cur = {...cur, r:NROWS-1, c};
  positionSel(); updateStatus();
}
function selectRow(r, add){
  selAnchor = {r, c:0};
  cur = {...cur, r, c:NCOLS-1};
  positionSel(); updateStatus();
}

/* update status bar sums */
function updateStatus(){
  const s = selRange();
  const n = (s.r2-s.r1+1)*(s.c2-s.c1+1);
  let sum=0, cnt=0;
  for(let r=s.r1;r<=s.r2;r++) for(let c=s.c1;c<=s.c2;c++){
    const v = evalCell(r,c);
    if(typeof v==="number"){ sum+=v; cnt++; }
  }
  $("#st-sel").textContent = n>1 ? `${cellRef(s.r1,s.c1)}:${cellRef(s.r2,s.c2)} · ${n} cells` : "Ready";
  $("#st-sum").textContent = cnt>1 ? `Sum: ${fmtNum(Math.round(sum*1e6)/1e6)}   Average: ${fmtNum(Math.round(sum/cnt*1e6)/1e6)}   Count: ${cnt}` : "";
}

/* context menu */
let ctxMenu = null;
function showCellMenu(x,y){
  hideCtx();
  const m = document.createElement("div");
  m.className="popover show";
  m.style.left=Math.min(x,innerWidth-220)+"px"; m.style.top=Math.min(y,innerHeight-260)+"px";
  const items = [
    ["Cut","cut","✂"],["Copy","copy","⧉"],["Paste","paste","📋"],
    ["","","sep"],
    ["Insert Cells…","insertcells","➕"],["Delete Cells…","deletecells","➖"],
    ["Insert Row","insrow","▭"],["Insert Column","inscol","▯"],
    ["","","sep"],
    ["Clear Contents","clearcontents","⌫"],["Format Cells…","fmtdialog","🖌"],
    ["","","sep"],
    ["New Comment","comment","💬"],["Data Validation…","datavalidation","☑"],
  ];
  m.innerHTML = items.map(([l,c,i])=>
    c==="sep" ? `<div class="psep"></div>` :
    `<button class="pitem" data-cmcmd="${c}"><span class="pi">${i}</span>${l}</button>`).join("");
  document.body.appendChild(m);
  ctxMenu = m;
}
function hideCtx(){ if(ctxMenu){ ctxMenu.remove(); ctxMenu=null; } }
document.addEventListener("click", e=>{
  const b = e.target.closest("[data-cmcmd]");
  if(b){ const cmd=b.dataset.cmcmd; hideCtx(); runCmd(cmd); }
  else if(!e.target.closest(".popover")) hideCtx();
});

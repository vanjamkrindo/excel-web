/* ============ 08 · Excel features: styles, conditional formatting, page setup ============ */
let HEADINGS_ON = true, FXBAR_ON = true;
let AUTOSAVE_ON = true;
let painter = null;            // format painter payload

/* ---------- sheet model helpers ---------- */
function blankSheet(name){
  return {name, cells:{}, cols:{}, rows:{}, merges:[], frozen:{rows:0,cols:0}, charts:[], filter:null,
    condRules:[], validations:[], hidden:false, tabColor:null, protect:false, view:"normal",
    pageSetup:{margins:"normal", orient:"portrait", size:"A4", printArea:null, fit:"auto", scale:100, gridPrint:false}};
}
function normalizeSheet(s){
  s.cells = s.cells||{}; s.cols = s.cols||{}; s.rows = s.rows||{};
  s.merges = s.merges||[]; s.frozen = s.frozen||{rows:0,cols:0}; s.charts = s.charts||[];
  s.condRules = s.condRules||[]; s.validations = s.validations||[];
  s.view = s.view||"normal";
  s.pageSetup = Object.assign({margins:"normal", orient:"portrait", size:"A4", printArea:null, fit:"auto", scale:100, gridPrint:false}, s.pageSetup||{});
  return s;
}

/* ---------- injected chrome CSS ---------- */
(function(){
  const st = document.createElement("style");
  st.id = "xl-extra-css";
  st.textContent = `
    #gridtable.nohead thead th, #gridtable.nohead tbody th{display:none !important;}
    body.nofx #fxbar{display:none !important;}
    #gridtable tr.pb-break td{border-bottom:2px dashed #4472c4 !important;}
    #gridtable td.databar{position:relative;}
    #gridtable td .dbar{position:absolute;left:0;bottom:0;height:3px;background:#638ec6;opacity:.85;}
    #gridtable td .cmt{position:absolute;top:0;right:0;width:0;height:0;border-top:5px solid #c00000;border-left:5px solid transparent;}
    #gridtable.painting td{cursor:copy;}
  `;
  document.head.appendChild(st);
})();

/* ---------- conditional formatting ---------- */
function cfFill(){ return "#ffc7ce"; }
function cfText(){ return "#9c0006"; }
function applyCondFormats(){
  const rules = sheet().condRules || [];
  if(!rules.length) return;
  const t = $("#gridtable");
  rules.forEach(rule=>{
    const vals = [];
    for(let r=rule.r1;r<=rule.r2;r++) for(let c=rule.c1;c<=rule.c2;c++) vals.push(evalCell(r,c));
    const nums = vals.filter(v=>typeof v==="number");
    let match = ()=>false;
    if(rule.type==="gt") match = v=> typeof v==="number" && v > rule.v1;
    else if(rule.type==="lt") match = v=> typeof v==="number" && v < rule.v1;
    else if(rule.type==="between") match = v=> typeof v==="number" && v >= Math.min(rule.v1,rule.v2) && v <= Math.max(rule.v1,rule.v2);
    else if(rule.type==="eq") match = v=> String(v)===String(rule.v1);
    else if(rule.type==="dupe"){
      const counts = {};
      vals.forEach(v=>{ const k=String(v); counts[k]=(counts[k]||0)+1; });
      match = v=> counts[String(v)]>1 && String(v)!=="";
    }
    else if(rule.type==="top"){
      const sorted = [...nums].sort((a,b)=>b-a);
      const cut = sorted[Math.max(0,(rule.v1||10)-1)];
      match = v=> typeof v==="number" && cut!==undefined && v>=cut;
    }
    else if(rule.type==="bottom"){
      const sorted = [...nums].sort((a,b)=>a-b);
      const cut = sorted[Math.max(0,(rule.v1||10)-1)];
      match = v=> typeof v==="number" && cut!==undefined && v<=cut;
    }
    else if(rule.type==="databar"){
      const mx = Math.max(1, ...nums.map(Math.abs));
      for(let r=rule.r1;r<=rule.r2;r++) for(let c=rule.c1;c<=rule.c2;c++){
        const td = t.querySelector(`td[data-r="${r}"][data-c="${c}"]`); if(!td) continue;
        const v = evalCell(r,c);
        if(typeof v!=="number") continue;
        const w = Math.round(Math.abs(v)/mx*100);
        td.style.backgroundImage =
          `linear-gradient(90deg, #638ec6 0%, #a9c4e8 ${w}%, rgba(0,0,0,0) ${w}%, rgba(0,0,0,0) 100%)`;
      }
      return;
    }
    else if(rule.type==="scale"){
      const mn = Math.min(...nums), mx = Math.max(...nums);
      for(let r=rule.r1;r<=rule.r2;r++) for(let c=rule.c1;c<=rule.c2;c++){
        const td = t.querySelector(`td[data-r="${r}"][data-c="${c}"]`); if(!td) continue;
        const v = evalCell(r,c);
        if(typeof v!=="number") continue;
        const p = mx===mn ? 0.5 : (v-mn)/(mx-mn);
        const rr = Math.round(248 - 158*p), gg = Math.round(105 + 100*p), bb = Math.round(107 + 20*p);
        td.style.background = `rgb(${rr},${gg},${bb})`;
        td.style.color = p>0.6 ? "#fff" : "#1b1b1b";
      }
      return;
    }
    for(let r=rule.r1;r<=rule.r2;r++) for(let c=rule.c1;c<=rule.c2;c++){
      if(!match(evalCell(r,c))) continue;
      const td = t.querySelector(`td[data-r="${r}"][data-c="${c}"]`); if(!td) continue;
      td.style.background = rule.fill || cfFill();
      td.style.color = rule.color || cfText();
    }
  });
}
function cfRuleDialog(type){
  const titles = {gt:"Greater Than", lt:"Less Than", between:"Between", eq:"Equal To", top:"Top 10 Items", bottom:"Bottom 10 Items"};
  const t = titles[type] || "Conditional Formatting";
  let body = `<div class="row"><label>Format cells that are</label><select id="cft">
    <option value="gt">greater than</option><option value="lt">less than</option>
    <option value="between">between</option><option value="eq">equal to</option>
    <option value="dupe">duplicate values</option></select></div>
    <div class="row"><label>Value</label><input type="text" id="cfv1" placeholder="e.g. 1000"></div>
    <div class="row"><label>and</label><input type="text" id="cfv2" placeholder="upper bound (Between only)"></div>
    <div class="row"><label>with</label><select id="cfp">
      <option value="red">Light Red Fill with Dark Red Text</option>
      <option value="yellow">Yellow Fill with Dark Yellow Text</option>
      <option value="green">Green Fill with Dark Green Text</option>
      <option value="blue">Light Blue Fill with Dark Blue Text</option></select></div>`;
  const d = dlg("dlg-cf", t+" — Conditional Formatting", body, dd=>{
    const s = selRange();
    const kind = dd.querySelector("#cft").value;
    const v1raw = dd.querySelector("#cfv1").value.trim();
    const v2raw = dd.querySelector("#cfv2").value.trim();
    const pal = {red:{fill:"#ffc7ce",color:"#9c0006"}, yellow:{fill:"#ffeb9c",color:"#9c6500"},
      green:{fill:"#c6efce",color:"#006100"}, blue:{fill:"#ddebf7",color:"#1f4e79"}}[dd.querySelector("#cfp").value];
    pushUndo("conditional format");
    sheet().condRules.push({...s, type:kind, v1:parseFloat(v1raw), v2:parseFloat(v2raw), fill:pal.fill, color:pal.color});
    renderGrid(); markDirty();
  }, "OK");
  const sel = d.querySelector("#cft");
  if(type && [...sel.options].some(o=>o.value===type)) sel.value = type;
  if(type==="top"||type==="bottom") d.querySelector("#cfv1").value = "10";
}
function cfQuick(type){
  const s = selRange();
  pushUndo("conditional format");
  if(type==="dupe"){
    sheet().condRules.push({...s, type:"dupe", fill:"#ffc7ce", color:"#9c0006"});
  } else if(type==="top" || type==="bottom"){
    sheet().condRules.push({...s, type, v1:10, fill:"#c6efce", color:"#006100"});
  } else {
    sheet().condRules.push({...s, type, fill:"#638ec6"});
  }
  renderGrid(); markDirty();
}
function cfManage(){
  const rules = sheet().condRules || [];
  dlg("dlg-cfman","Conditional Formatting Rules Manager",
    rules.length ? `<table style="font-size:12.5px;border-collapse:collapse;width:100%">
      <tr style="color:var(--muted);text-align:left"><th style="padding:4px 10px 4px 0">Range</th><th style="padding:4px">Rule</th><th style="padding:4px">Value</th><th></th></tr>
      ${rules.map((r,i)=>`<tr><td style="padding:4px 10px 4px 0">${cellRef(r.r1,r.c1)}:${cellRef(r.r2,r.c2)}</td>
        <td style="padding:4px">${r.type}</td><td style="padding:4px">${r.v1!==undefined&&!isNaN(r.v1)?r.v1:""}</td>
        <td style="padding:4px"><button class="btn" data-cfdel="${i}" style="padding:1px 8px">Delete</button></td></tr>`).join("")}
      </table>` : `<p style="font-size:12.5px;color:var(--muted)">No conditional formatting rules on this sheet.</p>`,
    ()=>false, "Close");
  const d = $("#dlg-cfman");
  d.addEventListener("click", e=>{
    const b = e.target.closest("[data-cfdel]"); if(!b) return;
    pushUndo("delete rule");
    sheet().condRules.splice(+b.dataset.cfdel,1);
    renderGrid(); markDirty(); d.close(); cfManage();
  });
}
function cfClear(){
  if(!(sheet().condRules||[]).length){ flashStatus("No rules to clear"); return; }
  pushUndo("clear rules");
  sheet().condRules = [];
  renderGrid(); markDirty(); flashStatus("All rules cleared from this sheet");
}

/* ---------- cell styles ---------- */
const CELL_STYLES = {
  normal:{bold:false, italic:false, underline:false, color:"#000000", bg:"#ffffff", size:11, numfmt:"General"},
  title:{bold:true, size:18, color:"#1f4e79", bg:"#ffffff"},
  h1:{bold:true, size:15, color:"#1f4e79"},
  h2:{bold:true, size:13, color:"#1f4e79"},
  h3:{bold:true, size:11, color:"#1f4e79"},
  h4:{bold:true, italic:true, size:11, color:"#1f4e79"},
  good:{bg:"#c6efce", color:"#006100"},
  bad:{bg:"#ffc7ce", color:"#9c0006"},
  neutral:{bg:"#ffeb9c", color:"#9c6500"},
  currency:{numfmt:"Currency"},
  percent:{numfmt:"Percent"},
  comma:{numfmt:"Number"},
  input:{bg:"#fff2cc", color:"#7f6000"},
  output:{bg:"#ddebf7", color:"#1f4e79"},
  calc:{bg:"#fce4d6", color:"#833c0c"},
  check:{border:15},
  accent1:{bg:"#4472c4", color:"#ffffff", bold:true},
  accent2:{bg:"#ed7d31", color:"#ffffff", bold:true},
  accent3:{bg:"#a5a5a5", color:"#ffffff", bold:true},
  accent4:{bg:"#ffc000", color:"#ffffff", bold:true},
  accent5:{bg:"#5b9bd5", color:"#ffffff", bold:true},
  accent6:{bg:"#70ad47", color:"#ffffff", bold:true},
};
function applyCellStyle(key){
  const preset = CELL_STYLES[key];
  if(!preset) return;
  pushUndo("cell style");
  eachSelCell((r,c)=>{
    const cl = getCellW(r,c);
    if(key==="normal"){ cl.s = {}; cl.numfmt = "General"; return; }
    cl.s = {...(cl.s||{}), ...preset};
    if(preset.numfmt){ cl.numfmt = preset.numfmt; delete cl.s.numfmt; }
  });
  recalcAll(); renderGrid(); markDirty();
}

/* ---------- format as table ---------- */
const TABLE_STYLES = {
  lightblue:{head:{bg:"#ddebf7", color:"#1f4e79"}, band:"#eef4fb"},
  lightgreen:{head:{bg:"#e2efda", color:"#375623"}, band:"#f1f8ed"},
  mediumblue:{head:{bg:"#4472c4", color:"#ffffff"}, band:"#d9e2f3"},
  mediumorange:{head:{bg:"#ed7d31", color:"#ffffff"}, band:"#fbe5d6"},
  darkgreen:{head:{bg:"#375623", color:"#ffffff"}, band:"#c6e0b4"},
};
function applyTableStyle(key){
  const s = selRange();
  if(s.r1===s.r2 && s.c1===s.c2){ flashStatus("Select a range with a header row first"); return; }
  const preset = TABLE_STYLES[key] || TABLE_STYLES.lightblue;
  pushUndo("format as table");
  for(let r=s.r1;r<=s.r2;r++){
    const isHead = r===s.r1;
    for(let c=s.c1;c<=s.c2;c++){
      const cl = getCellW(r,c);
      const st = {...(cl.s||{})};
      if(isHead){ st.bg = preset.head.bg; st.color = preset.head.color; st.bold = true; }
      else {
        st.bg = ((r-s.r1)%2===1) ? preset.band : "#ffffff";
        st.bold = false;
      }
      st.border = 15;
      cl.s = st;
    }
  }
  recalcAll(); renderGrid(); markDirty();
}

/* ---------- format painter ---------- */
function startFormatPainter(){
  const cl = getCell(cur.r,cur.c);
  painter = {style: JSON.parse(JSON.stringify((cl&&cl.s)||{})), numfmt: (cl&&cl.numfmt)||"General"};
  $("#gridtable").classList.add("painting");
  flashStatus("Format Painter: select the destination cells");
}
function applyPainter(r,c){
  if(!painter) return false;
  pushUndo("format painter");
  const s = selRange();
  for(let rr=s.r1;rr<=s.r2;rr++) for(let cc=s.c1;cc<=s.c2;cc++){
    const cl = getCellW(rr,cc);
    cl.s = JSON.parse(JSON.stringify(painter.style));
    cl.numfmt = painter.numfmt;
  }
  painter = null;
  $("#gridtable").classList.remove("painting");
  recalcAll(); renderGrid(); markDirty();
  flashStatus("Formatting applied");
  return true;
}

/* ---------- paste variants ---------- */
function pasteVariant(kind){
  if(!clipboard) { flashStatus("Nothing to paste"); return; }
  const saved = JSON.parse(JSON.stringify(clipboard));
  pushUndo("paste special");
  const base = selRange();
  for(let dr=0; dr<saved.h; dr++) for(let dc=0; dc<saved.w; dc++){
    const src = saved.cells[ck(dr,dc)];
    const r = base.r1+dr, c = base.c1+dc;
    if(r>=NROWS||c>=NCOLS||!src) continue;
    const dst = getCellW(r,c);
    if(kind==="values"){ if(src.v!==undefined) dst.v = src.v; else delete dst.v; delete dst.f; }
    if(kind==="formulas"){ if(src.f){ dst.f = shiftFormula(src.f, r-(saved.from.r1+dr), c-(saved.from.c1+dc)); } }
    if(kind==="formats"){ if(src.s) dst.s = JSON.parse(JSON.stringify(src.s)); if(src.numfmt) dst.numfmt = src.numfmt; }
  }
  recalcAll(); renderGrid(); markDirty();
}

/* ---------- fill down/right/up/left ---------- */
function fillDir(dir){
  const s = selRange();
  pushUndo("fill");
  const sh = sheet();
  if(dir==="down" || dir==="up"){
    if(s.r1===s.r2){ flashStatus("Select a range with more than one row"); return; }
    for(let c=s.c1;c<=s.c2;c++){
      const srcR = dir==="down" ? s.r1 : s.r2;
      const src = sh.cells[ck(srcR,c)];
      for(let r=s.r1;r<=s.r2;r++){
        if(r===srcR) continue;
        if(!src){ delete sh.cells[ck(r,c)]; continue; }
        const copy = JSON.parse(JSON.stringify(src));
        if(copy.f) copy.f = shiftFormula(copy.f, r-srcR, 0);
        sh.cells[ck(r,c)] = copy;
      }
    }
  } else {
    if(s.c1===s.c2){ flashStatus("Select a range with more than one column"); return; }
    for(let r=s.r1;r<=s.r2;r++){
      const srcC = dir==="right" ? s.c1 : s.c2;
      const src = sh.cells[ck(r,srcC)];
      for(let c=s.c1;c<=s.c2;c++){
        if(c===srcC) continue;
        if(!src){ delete sh.cells[ck(r,c)]; continue; }
        const copy = JSON.parse(JSON.stringify(src));
        if(copy.f) copy.f = shiftFormula(copy.f, 0, c-srcC);
        sh.cells[ck(r,c)] = copy;
      }
    }
  }
  recalcAll(); renderGrid(); markDirty();
}
function seriesDialog(){
  const s = selRange();
  dlg("dlg-series","Series",`
    <div class="row"><label>Series in</label><select id="sdir"><option value="down">Columns</option><option value="right">Rows</option></select></div>
    <div class="row"><label>Type</label><select id="stype"><option value="linear">Linear</option><option value="growth">Growth</option></select></div>
    <div class="row"><label>Step value</label><input type="number" id="sstep" value="1" step="any"></div>
    <div class="row"><label>Stop value</label><input type="number" id="sstop" value="" step="any" placeholder="optional"></div>`, d=>{
    const dir = d.querySelector("#sdir").value;
    const type = d.querySelector("#stype").value;
    const step = parseFloat(d.querySelector("#sstep").value)||1;
    const stopRaw = d.querySelector("#sstop").value;
    const stop = stopRaw==="" ? null : parseFloat(stopRaw);
    pushUndo("series");
    const sh = sheet();
    if(dir==="down"){
      let v = parseFloat(evalCell(s.r1,s.c1))||0;
      for(let r=s.r1;r<=s.r2;r++){
        sh.cells[ck(r,s.c1)] = {v:Math.round(v*1e10)/1e10};
        v = type==="growth" ? v*(1+step/100) : v+step;
        if(stop!==null && (step>0 ? v>stop : v<stop)) break;
      }
    } else {
      let v = parseFloat(evalCell(s.r1,s.c1))||0;
      for(let c=s.c1;c<=s.c2;c++){
        sh.cells[ck(s.r1,c)] = {v:Math.round(v*1e10)/1e10};
        v = type==="growth" ? v*(1+step/100) : v+step;
        if(stop!==null && (step>0 ? v>stop : v<stop)) break;
      }
    }
    recalcAll(); renderGrid(); markDirty();
  }, "OK");
}

/* ---------- clear variants ---------- */
function clearVariant(kind){
  pushUndo("clear");
  eachSelCell((r,c)=>{
    const sh = sheet(); const key = ck(r,c);
    if(!sh.cells[key]) return;
    if(kind==="contents"){ delete sh.cells[key].v; delete sh.cells[key].f; if(!Object.keys(sh.cells[key]).length) delete sh.cells[key]; }
    else if(kind==="formats"){ delete sh.cells[key].s; delete sh.cells[key].numfmt; if(!Object.keys(sh.cells[key]).length) delete sh.cells[key]; }
    else if(kind==="comments"){ delete sh.cells[key].comment; if(!Object.keys(sh.cells[key]).length) delete sh.cells[key]; }
    else delete sh.cells[key];
  });
  recalcAll(); renderGrid(); markDirty();
}

/* ---------- indents / double underline / fonts ---------- */
function indentBy(delta){
  const st = getStyle(cur.r,cur.c);
  const lvl = Math.max(0, Math.min(8, (st.indent||0)+delta));
  setStyleOnSel({indent:lvl});
}
function bumpFont(delta){
  const st = getStyle(cur.r,cur.c);
  const curSize = st.size || 11;
  const list = SIZE_LIST.map(Number);
  let nx = curSize;
  if(delta>0){ nx = list.find(s=>s>curSize) || 72; }
  else { nx = [...list].reverse().find(s=>s<curSize) || 8; }
  setStyleOnSel({size:nx});
}
function toggleDoubleUnderline(){
  const st = getStyle(cur.r,cur.c);
  if(st.dblunder) setStyleOnSel({dblunder:false, underline:true});
  else setStyleOnSel({dblunder:true, underline:false});
}
function applyBorderBits(bits){
  pushUndo("border");
  const s = selRange();
  eachSelCell((r,c)=>{
    const cl = getCellW(r,c); const st = cl.s || {};
    if(bits===0){ delete st.border; cl.s = st; return; }
    if(bits===16){
      let b = 0;
      if(r===s.r1) b|=1;
      if(c===s.c2) b|=2;
      if(r===s.r2) b|=4;
      if(c===s.c1) b|=8;
      st.border = b;
    } else if(bits===17){
      st.border = 15;
    } else {
      st.border = bits;
    }
    cl.s = st;
  });
  renderGrid(); markDirty();
}

/* ---------- merge variants ---------- */
function mergeMode(mode){
  const s = selRange();
  if(s.r1===s.r2 && s.c1===s.c2){ flashStatus("Select more than one cell to merge"); return; }
  pushUndo("merge");
  const sh = sheet();
  if(mode==="unmerge"){
    sh.merges = sh.merges.filter(m=>!(m.r1<=s.r2 && m.r2>=s.r1 && m.c1<=s.c2 && m.c2>=s.c1));
    renderGrid(); markDirty(); return;
  }
  sh.merges = sh.merges.filter(m=>!(m.r1>=s.r1&&m.r2<=s.r2&&m.c1>=s.c1&&m.c2<=s.c2));
  if(mode==="center" || mode==="cells"){
    sh.merges.push({...s});
    if(mode==="center") eachSelCell((r,c)=>{ const cl=getCellW(r,c); cl.s={...(cl.s||{}), align:"center"}; });
  } else if(mode==="across"){
    for(let r=s.r1;r<=s.r2;r++) sh.merges.push({r1:r, r2:r, c1:s.c1, c2:s.c2});
  }
  renderGrid(); markDirty();
}

/* ---------- page setup ---------- */
const PAGE_MARGINS = {normal:{t:"0.75in",b:"0.75in",l:"0.7in",r:"0.7in"},
  wide:{t:"1in",b:"1in",l:"1in",r:"1in"}, narrow:{t:"0.25in",b:"0.25in",l:"0.25in",r:"0.25in"}};
const PAGE_SIZES = {A4:"A4", Letter:"Letter", Legal:"Legal", A3:"A3"};
function applyPageSetup(){
  const ps = (sheet() && sheet().pageSetup) || {};
  let st = document.getElementById("print-setup");
  if(!st){ st = document.createElement("style"); st.id="print-setup"; document.head.appendChild(st); }
  const m = PAGE_MARGINS[ps.margins||"normal"] || PAGE_MARGINS.normal;
  const size = PAGE_SIZES[ps.size||"A4"] || "A4";
  st.textContent = `@page{size:${size} ${ps.orient||"portrait"}; margin:${m.t} ${m.r} ${m.b} ${m.l};}
  @media print{
    #gridtable{zoom:${(ps.scale||100)/100};}
    ${ps.gridPrint ? "#gridtable td, #gridtable th{border-color:#999 !important;}" : "#gridtable td{border-color:transparent !important;}"}
  }`;
}
function pageSetupDialog(){
  const ps = sheet().pageSetup;
  dlg("dlg-page","Page Setup",`
    <div class="row"><label>Orientation</label><select id="po">
      <option value="portrait"${ps.orient==="portrait"?" selected":""}>Portrait</option>
      <option value="landscape"${ps.orient==="landscape"?" selected":""}>Landscape</option></select></div>
    <div class="row"><label>Paper size</label><select id="psz">
      ${Object.keys(PAGE_SIZES).map(k=>`<option${ps.size===k?" selected":""}>${k}</option>`).join("")}</select></div>
    <div class="row"><label>Margins</label><select id="pm">
      <option value="normal"${ps.margins==="normal"?" selected":""}>Normal</option>
      <option value="wide"${ps.margins==="wide"?" selected":""}>Wide</option>
      <option value="narrow"${ps.margins==="narrow"?" selected":""}>Narrow</option></select></div>
    <div class="row"><label>Scale</label><input type="number" id="psc" min="10" max="400" value="${ps.scale||100}"> %</div>
    <div class="row"><label>Print</label><label style="display:flex;gap:5px;align-items:center"><input type="checkbox" id="pg" ${ps.gridPrint?"checked":""}> Gridlines</label></div>
    <div class="row"><label>Print area</label><span id="pareap" style="font-size:12px;color:var(--muted)">${ps.printArea?`${cellRef(ps.printArea.r1,ps.printArea.c1)}:${cellRef(ps.printArea.r2,ps.printArea.c2)}`:"Not set — prints the used range"}</span></div>`,
    d=>{
      pushUndo("page setup");
      ps.orient = d.querySelector("#po").value;
      ps.size = d.querySelector("#psz").value;
      ps.margins = d.querySelector("#pm").value;
      ps.scale = parseInt(d.querySelector("#psc").value,10)||100;
      ps.gridPrint = d.querySelector("#pg").checked;
      applyPageSetup(); markDirty();
    }, "OK");
}
function setMargins(kind){
  if(kind==="custom"){ pageSetupDialog(); return; }
  pushUndo("margins");
  sheet().pageSetup.margins = kind;
  applyPageSetup(); markDirty(); flashStatus("Margins: "+kind);
}
function setOrient(o){ pushUndo("orientation"); sheet().pageSetup.orient = o; applyPageSetup(); markDirty(); }
function setPageSize(s){ pushUndo("paper size"); sheet().pageSetup.size = s; applyPageSetup(); markDirty(); }
function setPrintArea(){
  const s = selRange();
  pushUndo("print area");
  sheet().pageSetup.printArea = {...s};
  applyPageSetup(); markDirty();
  flashStatus("Print area: "+cellRef(s.r1,s.c1)+":"+cellRef(s.r2,s.c2));
}
function clearPrintArea(){ pushUndo("print area"); sheet().pageSetup.printArea = null; applyPageSetup(); markDirty(); flashStatus("Print area cleared"); }
function setFit(v){
  pushUndo("scale to fit");
  sheet().pageSetup.fit = v;
  flashStatus(v==="auto" ? "Scale to fit: automatic" : `Scale to fit: ${v} page(s) wide`);
  markDirty();
}
function scaleDialog(){
  const ps = sheet().pageSetup;
  dlg("dlg-scale","Scale to Fit",`
    <div class="row"><label>Adjust to</label><input type="number" id="scv" min="10" max="400" value="${ps.scale||100}"> % normal size</div>
    <div class="row"><label>Fit to</label><input type="number" id="scw" min="1" max="20" value="${ps.fit!=="auto"?ps.fit:1}"> page(s) wide</div>`,
    d=>{
      pushUndo("scale");
      ps.scale = parseInt(d.querySelector("#scv").value,10)||100;
      ps.fit = d.querySelector("#scw").value || "auto";
      applyPageSetup(); markDirty();
    }, "OK");
}

/* ---------- view / chrome ---------- */
function applyChrome(){
  const t = $("#gridtable");
  if(t) t.classList.toggle("nohead", !HEADINGS_ON);
  document.body.classList.toggle("nofx", !FXBAR_ON);
}
function applyView(){
  const t = $("#gridtable");
  if(!t) return;
  $$("#gridtable tr.pb-break").forEach(tr=>tr.classList.remove("pb-break"));
  if(sheet().view!=="pagebreak") return;
  const perPage = 42;
  for(let r=perPage-1;r<NROWS;r+=perPage){
    const tr = t.querySelector(`tbody tr:nth-child(${r+1})`);
    if(tr) tr.classList.add("pb-break");
  }
}
function toggleGridprint(){
  pushUndo("print gridlines");
  const ps = sheet().pageSetup;
  ps.gridPrint = !ps.gridPrint;
  applyPageSetup(); markDirty();
  flashStatus("Print gridlines: "+(ps.gridPrint?"on":"off"));
}
function zoomDialog(){
  dlg("dlg-zoom","Zoom",`<div class="row"><label>Magnification</label><input type="number" id="zv" min="10" max="400" value="${ZOOM}"> %</div>
    <div class="row"><label>Presets</label><select id="zp">${[50,75,100,125,150,200].map(z=>`<option${z===ZOOM?" selected":""}>${z}%</option>`).join("")}</select></div>`,
    d=>{
      const v = parseInt(d.querySelector("#zv").value,10)||100;
      setZoom(v);
    }, "OK");
}
function zoomToSelection(){
  const s = selRange();
  const w = (s.c2-s.c1+1)*64, h = (s.r2-s.r1+1)*20;
  const vw = $("#gridwrap").clientWidth, vh = $("#gridwrap").clientHeight;
  const z = Math.max(25, Math.min(400, Math.floor(Math.min(vw/Math.max(w,1), vh/Math.max(h,1))*100)));
  setZoom(z);
}

/* ---------- data tools: text to columns, validation, custom sort ---------- */
function textToColumnsDialog(){
  dlg("dlg-t2c","Text to Columns",`
    <div class="row"><label>Delimiter</label><select id="t2d">
      <option value=",">Comma</option><option value=";">Semicolon</option>
      <option value="\\t">Tab</option><option value=" ">Space</option><option value="|">Pipe</option></select></div>
    <div class="row"><label>Range</label><span style="font-size:12px;color:var(--muted)">Splits the selected cells into the columns to their right</span></div>`,
    d=>{
      const delim = d.querySelector("#t2d").value.replace("\\t","\t");
      const s = selRange();
      pushUndo("text to columns");
      const sh = sheet();
      for(let r=s.r1;r<=s.r2;r++){
        for(let c=s.c1;c<=s.c2;c++){
          const cl = sh.cells[ck(r,c)];
          if(!cl || cl.f || cl.v==null) continue;
          const parts = String(cl.v).split(delim);
          if(parts.length<2) continue;
          cl.v = parts[0];
          parts.slice(1).forEach((p,i)=>{
            const cc = c+1+i;
            if(cc>=NCOLS) return;
            sh.cells[ck(r,cc)] = {v: (p!=="" && !isNaN(+p)) ? +p : p};
          });
        }
      }
      recalcAll(); renderGrid(); markDirty();
    }, "Finish");
}
function dataValidationDialog(){
  dlg("dlg-dv","Data Validation",`
    <div class="row"><label>Allow</label><select id="dvt">
      <option value="any">Any value</option><option value="number">Whole number</option>
      <option value="decimal">Decimal</option><option value="text">Text length</option>
      <option value="list">List</option></select></div>
    <div class="row"><label>Data</label><select id="dvo">
      <option value="between">between</option><option value="gt">greater than</option>
      <option value="lt">less than</option><option value="eq">equal to</option></select></div>
    <div class="row"><label>Minimum</label><input type="text" id="dv1" placeholder="e.g. 0"></div>
    <div class="row"><label>Maximum</label><input type="text" id="dv2" placeholder="e.g. 100"></div>
    <div class="row"><label>List values</label><input type="text" id="dvl" placeholder="Yes,No,Maybe"></div>
    <div class="row"><label>Error message</label><input type="text" id="dvm" value="The value you entered is not valid."></div>`,
    d=>{
      const s = selRange();
      pushUndo("data validation");
      const rule = {r1:s.r1,c1:s.c1,r2:s.r2,c2:s.c2,
        type: d.querySelector("#dvt").value, op: d.querySelector("#dvo").value,
        v1: d.querySelector("#dv1").value.trim(), v2: d.querySelector("#dv2").value.trim(),
        list: d.querySelector("#dvl").value.trim(), msg: d.querySelector("#dvm").value.trim()};
      sheet().validations = (sheet().validations||[]).filter(v=>!(v.r1===s.r1&&v.c1===s.c1&&v.r2===s.r2&&v.c2===s.c2));
      if(rule.type!=="any") sheet().validations.push(rule);
      renderGrid(); markDirty();
      flashStatus(rule.type==="any" ? "Validation cleared" : "Validation applied");
    }, "OK");
}
function validateInput(r,c,val){
  const rules = (sheet().validations||[]).filter(v=> r>=v.r1 && r<=v.r2 && c>=v.c1 && c<=v.c2);
  for(const v of rules){
    const str = String(val??"").trim();
    if(str==="") continue;
    if(v.type==="list"){
      const allowed = v.list.split(",").map(x=>x.trim()).filter(Boolean);
      if(allowed.length && !allowed.includes(str)){
        alert(v.msg || "Value must be one of: "+allowed.join(", "));
        return false;
      }
      continue;
    }
    const num = Number(str);
    if(isNaN(num)){ alert(v.msg || "Enter a number."); return false; }
    const n1 = parseFloat(v.v1), n2 = parseFloat(v.v2);
    if(v.type==="text"){
      if(v.op==="between" && (str.length<n1 || str.length>n2)){ alert(v.msg||"Text length must be between "+n1+" and "+n2+"."); return false; }
      if(v.op==="gt" && str.length<=n1){ alert(v.msg||"Text is too short."); return false; }
      if(v.op==="lt" && str.length>=n1){ alert(v.msg||"Text is too long."); return false; }
      continue;
    }
    if(v.type==="number" && !Number.isInteger(num)){ alert(v.msg||"Enter a whole number."); return false; }
    if(v.op==="between" && !isNaN(n1) && !isNaN(n2) && (num<n1 || num>n2)){ alert(v.msg||`Value must be between ${n1} and ${n2}.`); return false; }
    if(v.op==="gt" && !isNaN(n1) && num<=n1){ alert(v.msg||`Value must be greater than ${n1}.`); return false; }
    if(v.op==="lt" && !isNaN(n1) && num>=n1){ alert(v.msg||`Value must be less than ${n1}.`); return false; }
    if(v.op==="eq" && !isNaN(n1) && num!==n1){ alert(v.msg||`Value must equal ${n1}.`); return false; }
  }
  return true;
}
function customSortDialog(){
  const s = selRange();
  const cols = [];
  for(let c=s.c1;c<=s.c2;c++) cols.push(c);
  dlg("dlg-sort","Sort",`
    <div class="row"><label>Sort by</label><select id="sc">${cols.map(c=>`<option value="${c}">Column ${colName(c)}</option>`).join("")}</select></div>
    <div class="row"><label>Order</label><select id="so"><option value="asc">A to Z / Smallest to Largest</option><option value="desc">Z to A / Largest to Smallest</option></select></div>
    <div class="row"><label>Header row</label><input type="checkbox" id="sh" checked></div>`,
    d=>{
      const c = +d.querySelector("#sc").value;
      const asc = d.querySelector("#so").value==="asc";
      const hasHead = d.querySelector("#sh").checked;
      const sh = sheet();
      pushUndo("sort");
      const r1 = hasHead ? s.r1+1 : s.r1;
      const rows = [];
      for(let r=r1;r<=s.r2;r++){
        const row = [];
        for(let cc=0;cc<NCOLS;cc++) row.push(sh.cells[ck(r,cc)] ? JSON.parse(JSON.stringify(sh.cells[ck(r,cc)])) : null);
        rows.push(row);
      }
      const keyOf = row => { const cl = row[c]; return cl && !cl.f ? (cl.v ?? "") : ""; };
      rows.sort((a,b)=>{
        const x=keyOf(a), y=keyOf(b);
        if(typeof x==="number" && typeof y==="number") return asc? x-y : y-x;
        const sx=String(x).toLowerCase(), sy=String(y).toLowerCase();
        return asc ? sx.localeCompare(sy,"en") : sy.localeCompare(sx,"en");
      });
      rows.forEach((row,i)=>{ const r=r1+i; for(let cc=0;cc<NCOLS;cc++){ if(row[cc]) sh.cells[ck(r,cc)]=row[cc]; else delete sh.cells[ck(r,cc)]; }});
      recalcAll(); renderGrid(); markDirty();
    }, "OK");
}
function clearFilterCmd(){
  if(!sheet().filter){ flashStatus("No filter applied"); return; }
  sheet().filter = null; renderGrid(); flashStatus("Filter cleared");
}
function gotoSpecial(){
  dlg("dlg-gos","Go To Special",`
    <div class="row"><label>Select</label><select id="gs">
      <option value="formulas">Formulas</option>
      <option value="constants">Constants</option>
      <option value="blanks">Blanks (in used range)</option>
      <option value="comments">Comments</option>
      <option value="condfmt">Cells with conditional formatting</option></select></div>`,
    d=>{
      const mode = d.querySelector("#gs").value;
      const ur = usedRange();
      const hits = [];
      for(let r=0;r<=ur.r2;r++) for(let c=0;c<=ur.c2;c++){
        const cl = sheet().cells[ck(r,c)];
        if(mode==="formulas" && cl && cl.f) hits.push({r,c});
        if(mode==="constants" && cl && !cl.f && cl.v!==undefined) hits.push({r,c});
        if(mode==="blanks" && !cl) hits.push({r,c});
        if(mode==="comments" && cl && cl.comment) hits.push({r,c});
        if(mode==="condfmt" && (sheet().condRules||[]).some(x=>r>=x.r1&&r<=x.r2&&c>=x.c1&&c<=x.c2)) hits.push({r,c});
      }
      if(!hits.length){ flashStatus("No cells match that selection"); return; }
      cur = {...cur, r:hits[0].r, c:hits[0].c};
      selAnchor = {r:hits[0].r, c:hits[0].c};
      cur = {...cur, r:hits[hits.length-1].r, c:hits[hits.length-1].c};
      positionSel(); updateStatus();
      flashStatus(`${hits.length} cell(s) selected`);
    }, "OK");
}
function autofitColumns(){
  const s = selRange();
  pushUndo("autofit");
  for(let c=s.c1;c<=s.c2;c++){
    let w = 40;
    for(let r=0;r<NROWS;r++){
      const v = displayValue(r,c);
      if(!v) continue;
      w = Math.max(w, Math.min(320, v.length*7.2+14));
    }
    sheet().cols[c] = Math.round(w);
  }
  renderGrid(); markDirty();
}
function insertFunction(name){
  insertAtCursor(name+"(");
}
function freezeTop(){ pushUndo("freeze"); sheet().frozen = {rows:1,cols:0}; renderGrid(); }
function freezeLeft(){ pushUndo("freeze"); sheet().frozen = {rows:0,cols:1}; renderGrid(); }
function unfreeze(){ pushUndo("freeze"); sheet().frozen = {rows:0,cols:0}; renderGrid(); flashStatus("Panes unfrozen"); }
function setViewMode(mode){
  sheet().view = mode;
  applyView();
  flashStatus(mode==="pagebreak" ? "Page Break Preview" : "Normal view");
}
function helpDialog(){
  dlg("dlg-help","Help",`
    <p style="font-size:13px;margin-bottom:10px"><b>Excel (single-file edition)</b> — the ribbon, dialogs and grid mirror Microsoft Excel, running fully offline in one HTML file.</p>
    <table style="font-size:12.5px;border-collapse:collapse">
      ${[["Arrow keys / Tab / Enter","Move the selection"],
         ["F2 or any character","Edit the active cell"],
         ["Ctrl+C / Ctrl+X / Ctrl+V","Copy, cut, paste"],
         ["Ctrl+B / Ctrl+I / Ctrl+U","Bold, italic, underline"],
         ["Ctrl+1","Format Cells"],
         ["Ctrl+F","Find and Replace"],
         ["Ctrl+G","Go To"],
         ["Ctrl+S","Save as .xlsx"],
         ["F9","Recalculate all formulas"],
         ["Drag the fill handle","Copy or extend a series"],
         ["Drag a column/row header edge","Resize the column or row"]]
        .map(([k,v])=>`<tr><td style="padding:3px 16px 3px 0"><kbd>${k}</kbd></td><td style="padding:3px 0;color:var(--muted)">${v}</td></tr>`).join("")}
    </table>`, ()=>false, "Close");
}
function aboutDialog(){
  dlg("dlg-about","About",`
    <p style="font-size:13px;margin-bottom:8px"><b>Excel — single-file edition</b></p>
    <p style="font-size:12.5px;color:var(--muted);line-height:1.8">
    A spreadsheet engine in one HTML file: multi-sheet grid, formula engine (~70 functions), xlsx import/export,
    conditional formatting, cell styles, data validation, charts, print setup and a full Excel ribbon.<br>
    Bundled libraries: SheetJS (Apache-2.0). No network access, no install, no account.</p>`, ()=>false, "Close");
}

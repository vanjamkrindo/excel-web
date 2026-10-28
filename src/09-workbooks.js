/* ============ 09 · Workbook bar & multi-window ============ */
/* Two workbooks at once, as close to Android multi-window as a TWA allows:
   - a workbook switcher bar (every open book, by name, with a close button)
   - a side-by-side split view that shows two workbooks simultaneously
   - "Open in new window" launches a real OS window where the platform allows it */

let SPLIT_ON = false;      // side-by-side mode
let SPLIT_RIGHT = -1;      // index of the workbook in the right pane (-1 = none)
let SPLIT_RATIO = 0.5;     // fraction of width given to the left pane

/* ---------- workbook switcher bar ---------- */
function renderWorkbookBar(){
  const bar = $("#wbtabs");
  if(!bar) return;
  bar.innerHTML =
    WORKBOOKS.map((w,i)=>
      `<button class="wbtab${i===ACTIVE?" active":""}" data-wb="${i}" title="${esc(w.name)}">` +
        `<span class="wbnm">${esc(w.name)}</span>` +
        `<span class="wbdirty" style="display:${w.dirty?"inline":"none"}">•</span>` +
        `<span class="wbclose" data-wbclose="${i}" title="Close">✕</span>` +
      `</button>`
    ).join("") +
    `<button class="wbtab add" id="wbadd" title="New workbook">＋</button>` +
    `<button class="wbtab split" id="wbsplit" title="View side by side">▥</button>` +
    `<button class="wbtab newwin" id="wbnewwin" title="Open in new window">↗</button>`;
}

document.addEventListener("click", e=>{
  const cl = e.target.closest("[data-wbclose]");
  if(cl){ e.stopPropagation(); closeWB(+cl.dataset.wbclose); return; }
  const t = e.target.closest("[data-wb]");
  if(t){ switchWB(+t.dataset.wb); return; }
  if(e.target.id==="wbadd" || e.target.closest("#wbadd")){
    const w = createWorkbook();
    ACTIVE = WORKBOOKS.length-1; syncAliases(); renderAll(); afterChange();
    flashStatus(`${w.name} created`);
    return;
  }
  if(e.target.id==="wbsplit" || e.target.closest("#wbsplit")){ toggleSplit(); return; }
  if(e.target.id==="wbnewwin" || e.target.closest("#wbnewwin")){ openNewWindow(); return; }
});
document.addEventListener("dblclick", e=>{
  const t = e.target.closest("[data-wb]");
  if(t) renameWorkbook(+t.dataset.wb);
});

function renameWorkbook(i){
  const w = WORKBOOKS[i]; if(!w) return;
  const nm = prompt("Workbook name:", w.name);
  if(nm===null) return;
  w.name = uniqueName(nm.trim() || w.name, w.id);
  renderAll(); afterChange();
}

/* ---------- side-by-side split view ---------- */
function toggleSplit(){
  SPLIT_ON = !SPLIT_ON;
  if(SPLIT_ON){
    /* default the right pane to a different workbook if one exists */
    if(SPLIT_RIGHT<0 || SPLIT_RIGHT===ACTIVE || SPLIT_RIGHT>=WORKBOOKS.length){
      SPLIT_RIGHT = WORKBOOKS.length>1 ? (ACTIVE===0 ? 1 : 0) : -1;
    }
  }
  applySplit();
  flashStatus(SPLIT_ON ? "Side by side: on" : "Side by side: off");
}
function applySplit(){
  const wrap = $("#splitwrap");
  const app  = $("#app");
  if(!wrap || !app) return;
  app.classList.toggle("splitting", SPLIT_ON);
  wrap.style.display = SPLIT_ON ? "flex" : "none";
  const btn = $("#wbsplit"); if(btn) btn.classList.toggle("on", SPLIT_ON);
  if(SPLIT_ON){
    const l = $("#pane-left"), r = $("#pane-right");
    if(l) l.style.flex = `0 0 ${Math.round(SPLIT_RATIO*100)}%`;
    if(r) r.style.flex = "1 1 0";
    renderSplitGrids();
  }
  const sbtn = $("#wbsplit"); if(sbtn) sbtn.classList.toggle("on", SPLIT_ON);
}
function renderSplitLabels(){
  const lb = $("#split-left-name"), rb = $("#split-right-name");
  if(lb) lb.textContent = activeWB() ? activeWB().name : "";
  const rw = SPLIT_RIGHT>=0 ? WORKBOOKS[SPLIT_RIGHT] : null;
  if(rb) rb.textContent = rw ? rw.name : "(none)";
}
/* Dragging the divider resizes the two panes. */
function initSplitDivider(){
  const d = $("#split-divider");
  if(!d) return;
  let dragging = false;
  const move = ev=>{
    if(!dragging) return;
    const host = $("#splitwrap").getBoundingClientRect();
    const x = (ev.touches ? ev.touches[0].clientX : ev.clientX) - host.left;
    SPLIT_RATIO = Math.max(0.15, Math.min(0.85, x/host.width));
    applySplit();
  };
  const up = ()=>{ dragging = false; document.removeEventListener("mousemove", move); document.removeEventListener("mouseup", up);
                   document.removeEventListener("touchmove", move); document.removeEventListener("touchend", up); };
  const down = ev=>{ ev.preventDefault(); dragging = true;
    document.addEventListener("mousemove", move); document.addEventListener("mouseup", up);
    document.addEventListener("touchmove", move, {passive:false}); document.addEventListener("touchend", up); };
  d.addEventListener("mousedown", down);
  d.addEventListener("touchstart", down, {passive:false});
}

/* ---------- side-by-side grid rendering ---------- */
const SPLIT_ROWS = 60, SPLIT_COLS = 12;
function renderSplitGrids(){
  if(!SPLIT_ON) return;
  const panes = [$("#pane-left .splitgrid"), $("#pane-right .splitgrid")];
  const books = [ACTIVE, SPLIT_RIGHT];
  panes.forEach((tbl, pi)=>{
    if(!tbl) return;
    const bi = books[pi];
    const w = (bi>=0 && bi<WORKBOOKS.length) ? WORKBOOKS[bi] : null;
    const sh = w ? w.sheets[Math.min(w.cur?w.cur.sheet:0, w.sheets.length-1)] : null;
    if(!sh){ tbl.querySelector("thead").innerHTML=""; tbl.querySelector("tbody").innerHTML=""; return; }
    let head = "<tr><th style='width:38px'></th>";
    for(let c=0;c<SPLIT_COLS;c++) head += `<th>${colName(c)}</th>`;
    head += "</tr>";
    tbl.querySelector("thead").innerHTML = head;
    let body = "";
    for(let r=0;r<SPLIT_ROWS;r++){
      body += `<tr><th>${r+1}</th>`;
      for(let c=0;c<SPLIT_COLS;c++){
        const cell = sh.cells[ck(r,c)];
        let txt = "", al = "left";
        if(cell){
          if(cell.f){
            /* computed value, same as the main grid shows */
            txt = fmtNum(computeCell(w, r, c));
          }else if(cell.v!==undefined && cell.v!==""){
            txt = fmtNum(cell.v);
          }
          if(typeof txt === "number" || (txt!=="" && !isNaN(Number(txt)))) al = "right";
        }
        body += `<td style="text-align:${al}">${esc(txt)}</td>`;
      }
      body += "</tr>";
    }
    tbl.querySelector("tbody").innerHTML = body;
  });
  renderSplitLabels();
}

/* ---------- "open in new window" ---------- */
/* A TWA/APK has one activity, so this cannot start a second Android window.
   Where the platform allows it we open a real browser window (Chrome desktop,
   or Android's multi-window via the browser). Otherwise we offer the split view. */
function openNewWindow(){
  const w = activeWB();
  const url = new URL(location.href);
  url.searchParams.set("book", String(w ? w.id : 0));
  const win = window.open(url.href, "excelwb_" + (w ? w.id : 0), "noopener");
  if(!win){
    if(!SPLIT_ON){ toggleSplit(); flashStatus("Pop-up blocked — opened side by side instead"); }
    else flashStatus("Pop-up blocked — use side by side");
  }else{
    flashStatus(`"${w?w.name:"Workbook"}” opened in a new window`);
  }
}

/* ---------- touch: swipe the workbook bar on narrow screens ---------- */
function initWBTouchScroll(){
  const bar = $("#wbtabs"); if(!bar) return;
  let sx=0, sl=0, down=false;
  bar.addEventListener("touchstart", e=>{ down=true; sx=e.touches[0].clientX; sl=bar.scrollLeft; }, {passive:true});
  bar.addEventListener("touchmove", e=>{ if(!down) return; bar.scrollLeft = sl - (e.touches[0].clientX - sx); }, {passive:true});
  bar.addEventListener("touchend", ()=>{ down=false; }, {passive:true});
}

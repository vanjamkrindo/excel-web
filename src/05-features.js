/* ============ 05 · Commands ============ */
function setStyleOnSel(patch){
  pushUndo("format");
  eachSelCell((r,c)=>{ const cl = getCellW(r,c); cl.s = {...(cl.s||{}), ...patch}; });
  recalcAll(); renderGrid(); markDirty();
}
function toggleStyle(key){
  const st = getStyle(cur.r,cur.c);
  setStyleOnSel({[key]: !st[key]});
}
function clearSelContent(){
  pushUndo("clear");
  eachSelCell((r,c)=>{ delete sheet().cells[ck(r,c)]; });
  recalcAll(); renderGrid(); markDirty();
}
function setNumFmt(val){
  pushUndo("number format");
  eachSelCell((r,c)=>{ const cl=getCellW(r,c); cl.numfmt=val; });
  recalcAll(); renderGrid(); markDirty();
}

function runCmd(cmd, arg){
  if(editing && cmd!=="fxcancel") commitEdit(null);
  switch(cmd){
    /* ---- clipboard ---- */
    case "cut": doCopy(true); break;
    case "copy": doCopy(false); break;
    case "paste": doPaste(); break;
    case "pastevalues": pasteVariant("values"); break;
    case "pasteformulas": pasteVariant("formulas"); break;
    case "pasteformats": pasteVariant("formats"); break;
    case "formatpainter": startFormatPainter(); break;
    case "clear": case "clearcontents": clearVariant("contents"); break;
    case "cleardialog": clearVariant("contents"); break;
    case "clearall": clearVariant("all"); break;
    case "clearformats": clearVariant("formats"); break;
    case "clearcomments": clearVariant("comments"); break;
    /* ---- font ---- */
    case "bold": toggleStyle("bold"); break;
    case "italic": toggleStyle("italic"); break;
    case "underline": toggleStyle("underline"); break;
    case "dblunderline": toggleDoubleUnderline(); break;
    case "strike": toggleStyle("strike"); break;
    case "wrap": toggleStyle("wrap"); break;
    case "incfont": bumpFont(1); break;
    case "decfont": bumpFont(-1); break;
    case "fontname": setStyleOnSel({font:arg}); break;
    case "fontsize": setStyleOnSel({size:parseInt(arg,10)}); break;
    case "border": if(arg!==undefined) applyBorderBits(parseInt(arg,10)); else borderDialog(); break;
    case "borders": borderDialog(); break;
    /* ---- alignment ---- */
    case "alignleft": setStyleOnSel({align:"left"}); break;
    case "aligncenter": setStyleOnSel({align:"center"}); break;
    case "alignright": setStyleOnSel({align:"right"}); break;
    case "valigntop": setStyleOnSel({valign:"top"}); break;
    case "valignmiddle": setStyleOnSel({valign:"middle"}); break;
    case "valignbottom": setStyleOnSel({valign:"bottom"}); break;
    case "indinc": indentBy(1); break;
    case "inddec": indentBy(-1); break;
    case "mergecenter": mergeMode("center"); break;
    case "mergeacross": mergeMode("across"); break;
    case "mergecells": mergeMode("cells"); break;
    case "unmerge": mergeMode("unmerge"); break;
    case "merge": mergeMode("cells"); break;
    /* ---- number ---- */
    case "numfmt": setNumFmt(arg); break;
    case "percent": setNumFmt("Percent"); break;
    case "currency": setNumFmt("Currency"); break;
    case "comma": setNumFmt("Number"); break;
    case "decplus": adjustDecimals(1); break;
    case "decminus": adjustDecimals(-1); break;
    case "decdialog": decDialog(); break;
    /* ---- styles ---- */
    case "cfgt": cfRuleDialog("gt"); break;
    case "cflt": cfRuleDialog("lt"); break;
    case "cfbetween": cfRuleDialog("between"); break;
    case "cfeq": cfRuleDialog("eq"); break;
    case "cfdupe": cfQuick("dupe"); break;
    case "cftop": cfQuick("top"); break;
    case "cfbottom": cfQuick("bottom"); break;
    case "cfdatabar": cfQuick("databar"); break;
    case "cfscale": cfQuick("scale"); break;
    case "cfmanage": cfManage(); break;
    case "cfclear": cfClear(); break;
    case "tablestyle": if(arg) applyTableStyle(arg); else applyTableStyle("lightblue"); break;
    case "cellstyle": if(arg) applyCellStyle(arg); else applyCellStyle("h1"); break;
    /* ---- cells ---- */
    case "insertcells": insDelDialog("insert"); break;
    case "deletecells": insDelDialog("delete"); break;
    case "insrow": insertRows(); break;
    case "inscol": insertCols(); break;
    case "delrow": deleteRows(); break;
    case "delcol": deleteCols(); break;
    case "rowheight": promptNum("Row height (px):", rowH(cur.r), v=>{ pushUndo("row height"); const s=selRange(); for(let r=s.r1;r<=s.r2;r++) sheet().rows[r]=v; renderGrid(); markDirty(); }); break;
    case "colwidth": promptNum("Column width (px):", colW(cur.c), v=>{ pushUndo("column width"); const s=selRange(); for(let c=s.c1;c<=s.c2;c++) sheet().cols[c]=v; renderGrid(); markDirty(); }); break;
    case "autofit": autofitColumns(); break;
    case "fmtdialog": fmtDialog(); break;
    case "renamesheet": renameSheet(cur.sheet); break;
    case "hidesheet": wb.sheets[cur.sheet].hidden = true; renderSheets(); markDirty(); flashStatus("Sheet hidden"); break;
    case "unhidesheet": { const hid = wb.sheets.map((s,j)=>({s,j})).filter(x=>x.s.hidden);
      if(!hid.length){ flashStatus("No hidden sheets"); break; }
      dlg("dlg-unhide2","Unhide",`<div class="row"><label>Sheet</label><select id="uh2">${hid.map(h=>`<option value="${h.j}">${esc(h.s.name)}</option>`).join("")}</select></div>`, d=>{
        pushUndoAll(); wb.sheets[+d.querySelector("#uh2").value].hidden=false; renderSheets(); markDirty();
      }, "OK"); break; }
    /* ---- editing ---- */
    case "autosum": autoSum("SUM"); break;
    case "autosumfn": autoSum(arg||"SUM"); break;
    case "filldown": fillDir("down"); break;
    case "fillright": fillDir("right"); break;
    case "fillup": fillDir("up"); break;
    case "fillleft": fillDir("left"); break;
    case "series": seriesDialog(); break;
    case "sortasc": sortSel(true); break;
    case "sortdesc": sortSel(false); break;
    case "customsort": customSortDialog(); break;
    case "filter": toggleFilter(); break;
    case "clearfilter": clearFilterCmd(); break;
    case "finddialog": findDialog(); break;
    case "replacedialog": findDialog(); break;
    case "gotodialog": gotoDialog(); break;
    case "gotospecial": gotoSpecial(); break;
    /* ---- insert tab ---- */
    case "chartbar": insertChart("bar"); break;
    case "chartline": insertChart("line"); break;
    case "chartpie": insertChart("pie"); break;
    case "symdialog": symDialog(); break;
    /* ---- formulas ---- */
    case "fxdialog": fxDialog(); break;
    case "insertfn": insertFunction(arg||"SUM"); break;
    case "recalc": recalcAll(); renderGrid(); flashStatus("Workbook recalculated"); break;
    case "showformulas": SHOW_FORMULAS = !SHOW_FORMULAS; renderGrid(); break;
    case "calcopts": dlg("dlg-calc","Calculation Options",
      `<p style="font-size:12.5px;line-height:1.7">This workbook recalculates automatically after every change.<br>
      Use <kbd>F9</kbd> or <b>Calculate Now</b> to force a full recalculation.</p>`, ()=>false, "Close"); break;
    /* ---- page layout ---- */
    case "settheme": if(arg) setTheme(arg); else setTheme(document.documentElement.getAttribute("data-theme")==="dark"?"light":"dark"); break;
    case "setmargins": setMargins(arg||"normal"); break;
    case "setorient": setOrient(arg||"portrait"); break;
    case "orient": setOrient(arg||"portrait"); break;
    case "setpsize": setPageSize(arg||"A4"); break;
    case "setprintarea": setPrintArea(); break;
    case "clearprintarea": clearPrintArea(); break;
    case "setfit": setFit(arg||"auto"); break;
    case "pagesetup": pageSetupDialog(); break;
    case "scaledialog": scaleDialog(); break;
    case "gridlines": GRID_ON = !GRID_ON; renderGrid(); break;
    case "gridprint": toggleGridprint(); break;
    case "headings": HEADINGS_ON = !HEADINGS_ON; applyChrome(); break;
    case "formulabar": FXBAR_ON = !FXBAR_ON; applyChrome(); break;
    /* ---- data tools ---- */
    case "texttocolumns": textToColumnsDialog(); break;
    case "deldup": deleteDupes(); break;
    case "datavalidation": dataValidationDialog(); break;
    /* ---- review ---- */
    case "protect": toggleProtect(); break;
    case "comment": commentDialog(); break;
    case "delcomment": deleteComment(); break;
    case "stats": statsDialog(); break;
    case "spell": spellCheck(); break;
    /* ---- view ---- */
    case "freeze": case "freezedlg": freezeDialog(); break;
    case "freezetop": freezeTop(); break;
    case "freezeleft": freezeLeft(); break;
    case "unfreeze": unfreeze(); break;
    case "viewnormal": setViewMode("normal"); break;
    case "viewpagebreak": setViewMode("pagebreak"); break;
    case "zoom": setZoom(parseInt(arg,10)); break;
    case "zoomdialog": zoomDialog(); break;
    case "zoom100": setZoom(100); break;
    case "zoomsel": zoomToSelection(); break;
    case "theme": toggleTheme(); break;
    /* ---- help ---- */
    case "help": helpDialog(); break;
    case "about": aboutDialog(); break;
    /* ---- file ---- */
    case "save": saveBook(); break;
    case "saveas": downloadXlsx(); break;
    case "open": openPicker(); break;
    case "new": if(confirm("Create a new workbook? Unsaved changes are lost.")){ wb=newWorkbook(); cur={sheet:0,r:0,c:0}; selAnchor=null; afterChange(); renderSheets(); setDocTitle(); showTab("Home"); } break;
    case "print": printSheet(); break;
    default:
      if(MENUS[cmd]) openMenu(document.body, cmd);
      break;
  }
}
function adjustDecimals(delta){
  const curFmt = numFmtOf(cur.r,cur.c);
  const m = /^Dec(\d+)$/.exec(curFmt);
  const n = Math.max(0, Math.min(10, (m ? +m[1] : 2) + delta));
  setNumFmt(n===0 ? "Number" : "Dec"+n);
}
function setZoom(z){
  z = Math.max(25, Math.min(400, z||100));
  ZOOM = z;
  $("#gridtable").style.zoom = z/100;
  const sl = $("#zoomslider"), lb = $("#zoomlabel");
  if(sl) sl.value = Math.min(200, Math.max(50, z));
  if(lb) lb.textContent = z+"%";
}

/* ---- clipboard ---- */
function doCopy(cut){
  const s = selRange();
  const cells = {};
  for(let r=s.r1;r<=s.r2;r++) for(let c=s.c1;c<=s.c2;c++){
    const cl = getCell(r,c); if(cl) cells[ck(r-s.r1, c-s.c1)] = JSON.parse(JSON.stringify(cl));
  }
  clipboard = {cells, w:s.c2-s.c1+1, h:s.r2-s.r1+1, cut, from:{...selRange()}};
  if(navigator.clipboard && !cut){
    const lines=[];
    for(let r=s.r1;r<=s.r2;r++){ const L=[]; for(let c=s.c1;c<=s.c2;c++) L.push(rawOf(r,c)); lines.push(L.join("\t")); }
    navigator.clipboard.writeText(lines.join("\n")).catch(()=>{});
  }
  const dashed = $("#gridtable");
  flashStatus(cut?"Cut":"Copied");
}
function doPaste(){
  if(!clipboard) { flashStatus("Nothing to paste"); return; }
  pushUndo("paste");
  const sh = sheet();
  if(clipboard.cut){
    for(let r=clipboard.from.r1;r<=clipboard.from.r2;r++)
      for(let c=clipboard.from.c1;c<=clipboard.from.c2;c++) delete sh.cells[ck(r,c)];
  }
  const base = selRange();
  for(let dr=0; dr<clipboard.h; dr++)
    for(let dc=0; dc<clipboard.w; dc++){
      const src = clipboard.cells[ck(dr,dc)];
      const r = base.r1+dr, c = base.c1+dc;
      if(r>=NROWS||c>=NCOLS) continue;
      if(src){
        const copy = JSON.parse(JSON.stringify(src));
        if(copy.f && !clipboard.cut){
          const dR = r - (clipboard.from.r1+dr), dC = c - (clipboard.from.c1+dc);
          if(dR||dC) copy.f = shiftFormula(copy.f, dR, dC);
        }
        sh.cells[ck(r,c)] = copy;
      }
      else delete sh.cells[ck(r,c)];
    }
  if(clipboard.cut) clipboard = null;
  recalcAll(); renderGrid(); markDirty();
}

/* ---- insert/delete rows & cols ---- */
function shiftMapKeys(sh, mapFn){
  const nc = {};
  for(const k in sh.cells){
    const [r,c]=k.split(",").map(Number); const p = mapFn(r,c);
    if(p && p.r<NROWS && p.c<NCOLS) nc[ck(p.r,p.c)] = sh.cells[k];
  }
  sh.cells = nc;
}
function insertRows(){
  const s = selRange();
  pushUndo("insert rows");
  const sh = sheet(), n = s.r2-s.r1+1;
  shiftMapKeys(sh,(r,c)=> r>=s.r1 ? {r:r+n,c} : {r,c});
  sh.merges = sh.merges.map(m=> m.r1>=s.r1?{...m,r1:m.r1+n,r2:m.r2+n} : m);
  renderGrid(); markDirty();
}
function insertCols(){
  const s = selRange();
  pushUndo("insert columns");
  const sh = sheet(), n = s.c2-s.c1+1;
  shiftMapKeys(sh,(r,c)=> c>=s.c1 ? {r,c:c+n} : {r,c});
  sh.merges = sh.merges.map(m=> m.c1>=s.c1?{...m,c1:m.c1+n,c2:m.c2+n} : m);
  renderGrid(); markDirty();
}
function deleteRows(){
  const s = selRange();
  pushUndo("delete rows");
  const sh = sheet(), n = s.r2-s.r1+1;
  shiftMapKeys(sh,(r,c)=> r<s.r1 ? {r,c} : (r>=s.r1+n ? {r:r-n,c} : null));
  sh.merges = sh.merges.filter(m=> !(m.r1>=s.r1&&m.r2<=s.r2)).map(m=> m.r1>=s.r1?{...m,r1:m.r1-n,r2:m.r2-n}:m);
  renderGrid(); markDirty();
}
function deleteCols(){
  const s = selRange();
  pushUndo("delete columns");
  const sh = sheet(), n = s.c2-s.c1+1;
  shiftMapKeys(sh,(r,c)=> c<s.c1 ? {r,c} : (c>=s.c1+n ? {r,c:c-n} : null));
  sh.merges = sh.merges.filter(m=> !(m.c1>=s.c1&&m.c2<=s.c2)).map(m=> m.c1>=s.c1?{...m,c1:m.c1-n,c2:m.c2-n}:m);
  renderGrid(); markDirty();
}

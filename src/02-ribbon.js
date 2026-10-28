/* ============ 02 · Ribbon (Microsoft Excel layout) ============ */
/* ---------- icon set: inline SVG in Excel's line style ---------- */
const sv = (d, o) => `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.15" stroke-linecap="round" stroke-linejoin="round"${o? " " + o : ""}>${d}</svg>`;
/* 4 horizontal rules — the Excel alignment family */
const bars = (a,b,c,d) => `<path d="M2 3h12M${a} 6.6h${12-a+0}M2 10.2h12M${c} 13.8h${12-c}"/>`;
const alignL = sv(`<path d="M2 2.5h12M2 6.2h8M2 9.9h12M2 13.6h6"/>`);
const alignC = sv(`<path d="M2 2.5h12M4 6.2h8M2 9.9h12M5 13.6h6"/>`);
const alignR = sv(`<path d="M2 2.5h12M6 6.2h8M2 9.9h12M8 13.6h6"/>`);
const valT = sv(`<rect x="2.5" y="2.5" width="11" height="11" rx="1"/><path d="M5 5.2h6M5 8.2h6M5 11.2h6" opacity=".35"/><path d="M2.5 5.4h11" stroke-width="1.7"/>`);
const valM = sv(`<rect x="2.5" y="2.5" width="11" height="11" rx="1"/><path d="M5 5.2h6M5 8.2h6M5 11.2h6" opacity=".35"/><path d="M2.5 8.2h11" stroke-width="1.7"/>`);
const valB = sv(`<rect x="2.5" y="2.5" width="11" height="11" rx="1"/><path d="M5 5.2h6M5 8.2h6M5 11.2h6" opacity=".35"/><path d="M2.5 11h11" stroke-width="1.7"/>`);
const I = {
  /* clipboard */
  paste: sv(`<rect x="3" y="3" width="8.5" height="11" rx="1"/><path d="M5.5 3V1.8h3.5V3"/><rect x="6" y="6.4" width="7" height="7.6" rx="1" fill="currentColor" fill-opacity=".12"/><path d="M7.6 9.2h3.8M7.6 11.4h2.6"/>`),
  cut: sv(`<circle cx="4" cy="12" r="1.7"/><circle cx="11.6" cy="12" r="1.7"/><path d="M5.2 10.7L12 2.2M10.4 10.7L3.6 2.2"/>`),
  copy: sv(`<rect x="2.6" y="2.6" width="8" height="9.4" rx="1"/><rect x="5.6" y="5.2" width="8" height="9.2" rx="1" fill="#fff"/><path d="M7.6 8h4.4M7.6 10.2h4.4"/>`),
  brush: sv(`<path d="M4.2 9.6l5.2-5.2 1.6 1.6-5.2 5.2z"/><path d="M4.2 9.6c-1 .4-1.6 1.2-1.6 2.4 0 1 .6 1.6 1.6 1.6 1.2 0 2-.6 2.4-1.6"/><path d="M10.4 3.4l1.4-1.4 1.6 1.6-1.4 1.4"/>`),
  /* font */
  bold: "B", italic: "I", under: "U", strike: "S",
  fcolor: "A", bcolor: sv(`<path d="M4 8.4l4.4-4.4 3.6 3.6-4.4 4.4z"/><path d="M8.4 4l1.8-1.8 3.6 3.6-1.8 1.8"/><path d="M2.4 13.4h11.2"/>`),
  border: sv(`<rect x="2.4" y="2.4" width="11.2" height="11.2" rx="1"/><path d="M8 2.4v11.2M2.4 8h11.2" stroke-dasharray="1.6 1.6" opacity=".85"/>`),
  growf: sv(`<path d="M3.4 12l2.6-6 2.6 6M4.3 10h3.4"/><path d="M11.4 6.6l1.6-1.6 1.6 1.6M13 5v5.6"/>`),
  shrinkf: sv(`<path d="M3.4 12l2.6-6 2.6 6M4.3 10h3.4"/><path d="M11.4 5l1.6 1.6L14.6 5M13 10.6V5"/>`),
  /* alignment */
  al: alignL, ac: alignC, ar: alignR, at: valT, am: valM, ab: valB,
  orient: sv(`<rect x="2.6" y="2.6" width="10.8" height="10.8" rx="1"/><path d="M6 11l2.2-2.2M6 8.6V11h2.4"/>`),
  indinc: sv(`<path d="M2.5 3h11M2.5 8h11M2.5 13h11M6.4 6.2l1.8 1.8-1.8 1.8"/>`),
  inddec: sv(`<path d="M2.5 3h11M2.5 8h11M2.5 13h11M8.2 6.2L6.4 8l1.8 1.8"/>`),
  wrap: sv(`<path d="M2.6 3.4h10.8M2.6 7.6h6.2a2.2 2.2 0 010 4.4H6"/><path d="M7.4 10.4L5.8 12l1.6 1.6"/>`),
  merge: sv(`<rect x="2.4" y="4" width="11.2" height="8" rx="1"/><path d="M5 8h6M5.8 6.6L4.4 8l1.4 1.4M10.2 6.6L11.6 8l-1.4 1.4"/>`),
  /* number */
  pct: "%", cur: "$", comma: ",",
  decp: `<span style="font-size:8px;letter-spacing:-.4px">.00→</span>`,
  decm: `<span style="font-size:8px;letter-spacing:-.4px">←.0</span>`,
  /* styles */
  cond: sv(`<rect x="2.4" y="2.4" width="11.2" height="11.2" rx="1"/><path d="M2.4 6.2h11.2M2.4 9.9h11.2"/><path d="M2.4 6.2h5.6v7.4" fill="currentColor" fill-opacity=".18"/>`),
  table: sv(`<rect x="2.4" y="2.4" width="11.2" height="11.2" rx="1"/><path d="M2.4 6h11.2M6.2 6v7.6M9.9 6v7.6"/><path d="M2.4 6h11.2v-3.6H2.4z" fill="currentColor" fill-opacity=".18"/>`),
  cstyle: sv(`<rect x="2.2" y="3.4" width="11.6" height="9.6" rx=".8"/><path d="M2.2 6.6h11.6M6 6.6v6.4M9.8 6.6v6.4"/><path d="M2.2 6.6h11.6v-3.2H2.2z" fill="currentColor" fill-opacity=".22"/><path d="M3.8 4.6h1.2M6.2 4.6h1.2M8.6 4.6h1.2"/>`),
  /* cells */
  ins: sv(`<rect x="1.6" y="6" width="6.4" height="6.4" rx=".8"/><path d="M1.6 9.2h6.4M4.8 6v6.4"/><path d="M11 4.4h3.2M12.6 2.8v3.2" stroke-width="1.5"/><path d="M9.2 4.4h1.4" opacity="0"/>`),
  del: sv(`<rect x="1.6" y="6" width="6.4" height="6.4" rx=".8"/><path d="M1.6 9.2h6.4M4.8 6v6.4"/><path d="M10.6 3.2l3 3M13.6 3.2l-3 3" stroke-width="1.5"/>`),
  fmt: sv(`<rect x="1.6" y="6" width="6.4" height="6.4" rx=".8"/><path d="M1.6 9.2h6.4M4.8 6v6.4"/><path d="M9.6 4.2l3.4-1.6 1 2.2-3.4 1.6z"/><path d="M12 4.6l.8 2.6" stroke-width="1.3"/>`),
  /* editing */
  sum: "Σ",
  fill: sv(`<path d="M3 8.2l4.6-4.6 4 4-4.6 4.6z"/><path d="M6.2 3l1-1 4.6 4.6-1 1"/><path d="M13 9.6c.7 1 1.2 1.8 1.2 2.4a1.2 1.2 0 11-2.4 0c0-.6.5-1.4 1.2-2.4z"/>`),
  clear: sv(`<path d="M5.4 12.6h8M3.2 9.4l4.6-4.6 3.4 3.4-4.6 4.6z"/><path d="M5.6 12.2l-1.2 1.2H3"/><path d="M7.8 4.8l1.4-1.4 3.4 3.4-1.4 1.4"/>`),
  sortfilter: sv(`<path d="M2.6 3.4h6M2.6 6.4h4.4M2.6 9.4h2.8"/><path d="M11 3v9M9 10l2 2 2-2"/>`),
  find: sv(`<circle cx="7" cy="7" r="4.2"/><path d="M10.2 10.2l3.2 3.2"/>`),
  /* insert tab */
  table2: sv(`<rect x="2.4" y="2.4" width="11.2" height="11.2" rx="1"/><path d="M2.4 6h11.2M6.2 6v7.6M9.9 6v7.6"/>`),
  chart: sv(`<path d="M2.6 2.6v10.8h10.8"/><rect x="4.6" y="8.4" width="2" height="4" fill="currentColor" fill-opacity=".2"/><rect x="7.6" y="5.6" width="2" height="6.8" fill="currentColor" fill-opacity=".2"/><rect x="10.6" y="3.4" width="2" height="9" fill="currentColor" fill-opacity=".2"/>`),
  line: sv(`<path d="M2.6 2.6v10.8h10.8"/><path d="M4.4 10.6l3-3.2 2.4 1.8 3-3.6"/>`),
  pie: sv(`<circle cx="8" cy="8" r="5.4"/><path d="M8 2.6V8h5.4"/>`),
  pic: sv(`<rect x="2.4" y="3" width="11.2" height="10" rx="1"/><circle cx="6" cy="6.4" r="1.1"/><path d="M2.8 12l3.4-3.4 2.2 2.2 2-2 2.8 2.8"/>`),
  symbol: "Ω",
  /* page layout */
  theme: sv(`<circle cx="8" cy="8" r="5.4"/><path d="M8 2.6v10.8M2.6 8h10.8"/>`),
  margins: sv(`<rect x="2.4" y="3" width="11.2" height="10" rx="1" stroke-dasharray="1.8 1.6"/><path d="M5 5.6h6M5 8h6M5 10.4h4"/>`),
  size: sv(`<path d="M3.4 3h9.2M8 3v10M6 13h4"/>`),
  printarea: sv(`<rect x="2.6" y="2.6" width="10.8" height="10.8" rx="1" stroke-dasharray="1.8 1.6"/><path d="M5.4 8h5.2M8 5.4v5.2"/>`),
  breaks: sv(`<path d="M2.6 4.6h10.8" stroke-dasharray="1.6 1.6"/><path d="M2.6 8.2h10.8M2.6 11.8h10.8"/>`),
  scale: sv(`<rect x="2.6" y="2.6" width="10.8" height="10.8" rx="1"/><path d="M5.4 10.6l5.2-5.2M6.2 5.4h4.4v4.4"/>`),
  grid: sv(`<rect x="2.4" y="2.4" width="11.2" height="11.2" rx="1"/><path d="M6 2.4v11.2M10 2.4v11.2M2.4 6h11.2M2.4 10h11.2" opacity=".7"/>`),
  gridprint: sv(`<rect x="2.4" y="2.4" width="11.2" height="11.2" rx="1"/><path d="M6 2.4v11.2M10 2.4v11.2M2.4 6h11.2M2.4 10h11.2" opacity=".7"/><path d="M4 12.4l8-8" stroke-width="1.5"/>`),
  headings: sv(`<path d="M2.6 4h3.2M2.6 8h3.2M2.6 12h3.2M9.4 4h4M9.4 8h4M9.4 12h4" opacity=".8"/>`),
  orient: sv(`<rect x="2.6" y="2.6" width="10.8" height="10.8" rx="1"/><path d="M6 11l2.2-2.2M6 8.6V11h2.4"/>`),
  /* formulas tab */
  fx: "ƒx",
  catfin: sv(`<rect x="2.4" y="4.6" width="11.2" height="7" rx="1"/><path d="M5.4 4.6V3.4h5.2v1.2M5.4 8h5.2M8 6.6v2.8"/>`),
  catlog: sv(`<path d="M8 2.4v11.2M3.2 8h9.6"/><circle cx="8" cy="8" r="1.6"/>`),
  cattext: "T",
  catdate: sv(`<rect x="2.6" y="3.4" width="10.8" height="10" rx="1"/><path d="M2.6 6.4h10.8M5.4 2.4v2M10.6 2.4v2"/>`),
  catlookup: sv(`<circle cx="7" cy="7" r="4"/><path d="M10 10l3.4 3.4"/><path d="M5.2 7h3.6"/>`),
  catmath: "∑",
  recalc: sv(`<path d="M13 8a5 5 0 11-1.6-3.7"/><path d="M13.2 2.6v2.8h-2.8"/>`),
  audit: sv(`<rect x="2.4" y="2.4" width="11.2" height="11.2" rx="1"/><path d="M6.2 6.2h3.6v3.6H6.2z"/>`),
  /* data tab */
  filter: sv(`<path d="M2.6 3.4h10.8l-4.2 5v4.2l-2.4 1.2V8.4z"/>`),
  textcol: sv(`<path d="M3 8h6M6.6 5.6L3 8l3.6 2.4M13 8H9.4M10.6 5.6L13 8l-2.4 2.4" opacity=".9"/>`),
  dup: sv(`<rect x="2.4" y="2.4" width="8" height="8" rx="1"/><rect x="5.6" y="5.6" width="8" height="8" rx="1" fill="#fff"/>`),
  valid: sv(`<rect x="2.4" y="2.4" width="11.2" height="11.2" rx="1"/><path d="M4.8 8.2l2.2 2.2 4.2-4.6"/>`),
  subtotal: "Σ",
  /* review tab */
  spell: sv(`<path d="M2.8 8.2l3 3 7.4-7.4"/>`),
  stats: sv(`<path d="M2.6 2.6v10.8h10.8"/><rect x="4.6" y="8.4" width="2" height="4" fill="currentColor" fill-opacity=".2"/><rect x="7.6" y="5.6" width="2" height="6.8" fill="currentColor" fill-opacity=".2"/><rect x="10.6" y="3.4" width="2" height="9" fill="currentColor" fill-opacity=".2"/>`),
  comment: sv(`<path d="M2.6 3.4h10.8v7.2H7.4L4.4 13V10.6H2.6z"/>`),
  delcomment: sv(`<path d="M2.6 3.4h10.8v7.2H7.4L4.4 13V10.6H2.6z"/><path d="M6 6l4 3M10 6l-4 3"/>`),
  protect: sv(`<rect x="3.4" y="6.6" width="9.2" height="7" rx="1"/><path d="M5.6 6.6V4.8a2.4 2.4 0 014.8 0v1.8"/>`),
  lock: sv(`<rect x="3.4" y="6.6" width="9.2" height="7" rx="1"/><path d="M5.6 6.6V4.8a2.4 2.4 0 014.8 0"/>`),
  /* view tab */
  normal: sv(`<rect x="2.4" y="2.4" width="11.2" height="11.2" rx="1"/><path d="M2.4 5.6h11.2"/>`),
  layout: sv(`<rect x="2.4" y="2.4" width="11.2" height="11.2" rx="1"/><path d="M4.8 4.8h6.4v6.4H4.8z" opacity=".6"/>`),
  pagebreak: sv(`<rect x="2.4" y="2.4" width="11.2" height="11.2" rx="1"/><path d="M2.4 8h11.2" stroke-dasharray="1.6 1.6"/>`),
  formulabar: sv(`<rect x="2.4" y="3.4" width="11.2" height="9.2" rx="1"/><path d="M2.4 6.4h11.2M6.4 6.4v6.2"/>`),
  zoom: sv(`<circle cx="7" cy="7" r="4.2"/><path d="M10.2 10.2l3.2 3.2M5.2 7h3.6M7 5.2v3.6"/>`),
  zoom100: "100",
  zoomsel: sv(`<circle cx="7" cy="7" r="4.2"/><path d="M10.2 10.2l3.2 3.2"/>`),
  freeze: sv(`<rect x="2.4" y="2.4" width="11.2" height="11.2" rx="1"/><path d="M2.4 6.4h11.2M6.4 2.4v11.2" stroke-width="1.5"/><path d="M2.4 2.4h4v4h-4z" fill="currentColor" fill-opacity=".15"/>`),
  split: sv(`<rect x="2.4" y="2.4" width="11.2" height="11.2" rx="1"/><path d="M2.4 8h11.2M8 2.4v11.2" stroke-dasharray="1.6 1.6"/>`),
  /* misc */
  help: "?", macro: sv(`<path d="M4.4 3.4l8 4.6-8 4.6z"/>`), about: sv(`<circle cx="8" cy="8" r="5.4"/><path d="M8 7.2v4M8 5.2v.6"/>`),
};
/* Excel's big buttons are colour-coded — blue for structure, red for delete, orange for painter */
const ICOL = {
  paste:"#185abd", cut:"#185abd", copy:"#185abd", brush:"#c55a11", bcolor:"#185abd",
  cond:"#185abd", table:"#185abd", cstyle:"#185abd", ins:"#185abd", del:"#c00000", fmt:"#185abd",
  find:"#185abd", sortfilter:"#185abd", fill:"#185abd", clear:"#c00000", valid:"#185abd",
  comment:"#185abd", protect:"#c00000", freeze:"#185abd", grid:"#185abd", gridprint:"#185abd",
  chart:"#185abd", line:"#185abd", pie:"#185abd", pic:"#185abd", table2:"#185abd",
  catdate:"#185abd", catlookup:"#185abd", dup:"#185abd", textcol:"#185abd", filter:"#185abd",
};
for(const k in ICOL) if(typeof I[k] === "string") I[k] = I[k].replace(/currentColor/g, ICOL[k]);

const FONT_LIST = ["Calibri","Calibri Light","Arial","Arial Black","Times New Roman","Courier New","Georgia","Verdana","Tahoma","Segoe UI","Trebuchet MS","Comic Sans MS"];
const SIZE_LIST = ["8","9","10","11","12","14","16","18","20","22","24","26","28","36","48","72"];
const NUMFMTS = {
  General:{label:"General"}, Number:{label:"Number"}, Currency:{label:"Currency"}, Accounting:{label:"Accounting"},
  ShortDate:{label:"Short Date"}, LongDate:{label:"Long Date"}, Time:{label:"Time"},
  Percent:{label:"Percentage"}, Scientific:{label:"Scientific"}, Text:{label:"Text"},
};

/* dropdown menus for split buttons */
const MENUS = {
  paste:[
    {l:"Paste", c:"paste", i:"📋"},
    {l:"Paste Values", c:"pastevalues", i:"123"},
    {l:"Paste Formulas", c:"pasteformulas", i:"ƒx"},
    {l:"Paste Formatting", c:"pasteformats", i:"🖌"},
    {sep:1},
    {l:"Paste Special…", c:"fmtdialog", i:"⋯"},
  ],
  underline:[{l:"Underline", c:"underline", i:"U"},{l:"Double Underline", c:"dblunderline", i:"U̿"}],
  borders:[
    {l:"Bottom Border", c:"border", a:4, i:"▁"},{l:"Top Border", c:"border", a:1, i:"▔"},
    {l:"Left Border", c:"border", a:8, i:"▏"},{l:"Right Border", c:"border", a:2, i:"▕"},
    {sep:1},{l:"No Border", c:"border", a:0, i:"⬚"},{l:"All Borders", c:"border", a:15, i:"▦"},
    {l:"Outside Borders", c:"border", a:16, i:"⬜"},{l:"Thick Box Border", c:"border", a:17, i:"▣"},
  ],
  merge:[
    {l:"Merge & Center", c:"mergecenter", i:"⧉"},{l:"Merge Across", c:"mergeacross", i:"⧉"},
    {l:"Merge Cells", c:"mergecells", i:"⧉"},{sep:1},{l:"Unmerge Cells", c:"unmerge", i:"⧈"},
  ],
  condfmt:[
    {l:"Highlight Cells Greater Than…", c:"cfgt", i:"▲"},{l:"Highlight Cells Less Than…", c:"cflt", i:"▼"},
    {l:"Highlight Cells Between…", c:"cfbetween", i:"↔"},{l:"Highlight Duplicate Values", c:"cfdupe", i:"⧉"},
    {sep:1},{l:"Top 10 Items", c:"cftop", i:"⬆"},{l:"Bottom 10 Items", c:"cfbottom", i:"⬇"},
    {l:"Data Bars", c:"cfdatabar", i:"▮"},{l:"Color Scales", c:"cfscale", i:"🌈"},
    {sep:1},{l:"Manage Rules…", c:"cfmanage", i:"⋯"},{l:"Clear Rules", c:"cfclear", i:"✕"},
  ],
  astable:[
    {l:"Light Blue, Table Style", c:"tablestyle", a:"lightblue", i:"▦"},
    {l:"Light Green, Table Style", c:"tablestyle", a:"lightgreen", i:"▦"},
    {l:"Medium Blue, Table Style", c:"tablestyle", a:"mediumblue", i:"▦"},
    {l:"Medium Orange, Table Style", c:"tablestyle", a:"mediumorange", i:"▦"},
    {l:"Dark Green, Table Style", c:"tablestyle", a:"darkgreen", i:"▦"},
  ],
  cellstyles:[
    {l:"Normal", c:"cellstyle", a:"normal", i:"A"},
    {l:"Title", c:"cellstyle", a:"title", i:"T"},
    {l:"Heading 1", c:"cellstyle", a:"h1", i:"H1"},
    {l:"Heading 2", c:"cellstyle", a:"h2", i:"H2"},
    {l:"Heading 3", c:"cellstyle", a:"h3", i:"H3"},
    {l:"Heading 4", c:"cellstyle", a:"h4", i:"H4"},
    {sep:1},
    {l:"Good", c:"cellstyle", a:"good", i:"✓"},{l:"Bad", c:"cellstyle", a:"bad", i:"✕"},
    {l:"Neutral", c:"cellstyle", a:"neutral", i:"!"},
    {sep:1},
    {l:"Currency", c:"cellstyle", a:"currency", i:"$"},{l:"Percent", c:"cellstyle", a:"percent", i:"%"},
    {l:"Comma", c:"cellstyle", a:"comma", i:","},{l:"Input", c:"cellstyle", a:"input", i:"✎"},
    {l:"Output", c:"cellstyle", a:"output", i:"↗"},{l:"Calculation", c:"cellstyle", a:"calc", i:"Σ"},
    {l:"Check Cell", c:"cellstyle", a:"check", i:"☑"},
    {sep:1},
    {l:"Accent1", c:"cellstyle", a:"accent1", i:"1"},{l:"Accent2", c:"cellstyle", a:"accent2", i:"2"},
    {l:"Accent3", c:"cellstyle", a:"accent3", i:"3"},{l:"Accent4", c:"cellstyle", a:"accent4", i:"4"},
    {l:"Accent5", c:"cellstyle", a:"accent5", i:"5"},{l:"Accent6", c:"cellstyle", a:"accent6", i:"6"},
  ],
  insertcells:[{l:"Insert Cells…", c:"insertcells", i:"➕"},{l:"Insert Sheet Rows", c:"insrow", i:"▭"},{l:"Insert Sheet Columns", c:"inscol", i:"▯"}],
  deletecells:[{l:"Delete Cells…", c:"deletecells", i:"➖"},{l:"Delete Sheet Rows", c:"delrow", i:"▭"},{l:"Delete Sheet Columns", c:"delcol", i:"▯"}],
  format:[
    {l:"Row Height…", c:"rowheight", i:"↕"},{l:"Column Width…", c:"colwidth", i:"↔"},
    {l:"AutoFit Column Width", c:"autofit", i:"↔"},
    {sep:1},{l:"Format Cells…", c:"fmtdialog", i:"⋯"},{l:"Rename Sheet…", c:"renamesheet", i:"✎"},
    {l:"Hide Sheet", c:"hidesheet", i:"👁"},{l:"Unhide Sheet", c:"unhidesheet", i:"👁"},
    {sep:1},{l:"Protect Sheet", c:"protect", i:"🔒"},
  ],
  autosum:[
    {l:"Sum", c:"autosumfn", a:"SUM", i:"Σ"},{l:"Average", c:"autosumfn", a:"AVERAGE", i:"x̄"},
    {l:"Count Numbers", c:"autosumfn", a:"COUNT", i:"#"},{l:"Max", c:"autosumfn", a:"MAX", i:"▲"},
    {l:"Min", c:"autosumfn", a:"MIN", i:"▼"},
  ],
  fill:[
    {l:"Down", c:"filldown", i:"⬇"},{l:"Right", c:"fillright", i:"➡"},
    {l:"Up", c:"fillup", i:"⬆"},{l:"Left", c:"fillleft", i:"⬅"},
    {sep:1},{l:"Series…", c:"series", i:"⋯"},
  ],
  clear:[
    {l:"Clear All", c:"clearall", i:"✕"},{l:"Clear Formats", c:"clearformats", i:"🖌"},
    {l:"Clear Contents", c:"clearcontents", i:"⌫"},{l:"Clear Comments", c:"clearcomments", i:"💬"},
  ],
  sortfilter:[
    {l:"Sort A to Z", c:"sortasc", i:"⬆"},{l:"Sort Z to A", c:"sortdesc", i:"⬇"},
    {l:"Custom Sort…", c:"customsort", i:"⇅"},
    {sep:1},{l:"Filter", c:"filter", i:"⏷"},{l:"Clear Filter", c:"clearfilter", i:"✕"},
  ],
  findsel:[
    {l:"Find…", c:"finddialog", i:"🔍"},{l:"Replace…", c:"replacedialog", i:"⇄"},
    {l:"Go To…", c:"gotodialog", i:"→"},{l:"Go To Special…", c:"gotospecial", i:"⋯"},
  ],
  themes:[{l:"Office (Light)", c:"settheme", a:"light", i:"☀"},{l:"Dark Gray", c:"settheme", a:"dark", i:"🌙"}],
  margins:[
    {l:"Normal", c:"setmargins", a:"normal", i:"⬚"},{l:"Wide", c:"setmargins", a:"wide", i:"⬚"},
    {l:"Narrow", c:"setmargins", a:"narrow", i:"⬚"},{sep:1},{l:"Custom Margins…", c:"setmargins", a:"custom", i:"⋯"},
  ],
  orient:[{l:"Portrait", c:"setorient", a:"portrait", i:"▯"},{l:"Landscape", c:"setorient", a:"landscape", i:"▭"}],
  psize:[{l:"A4", c:"setpsize", a:"A4", i:"📄"},{l:"Letter", c:"setpsize", a:"Letter", i:"📄"},{l:"Legal", c:"setpsize", a:"Legal", i:"📄"},{l:"A3", c:"setpsize", a:"A3", i:"📄"}],
  printarea:[{l:"Set Print Area", c:"setprintarea", i:"▣"},{l:"Clear Print Area", c:"clearprintarea", i:"✕"}],
  widthfit:[{l:"1 page", c:"setfit", a:"1", i:"1"},{l:"2 pages", c:"setfit", a:"2", i:"2"},{l:"Automatic", c:"setfit", a:"auto", i:"A"}],
  freeze:[{l:"Freeze Panes", c:"freezedlg", i:"❄"},{l:"Freeze Top Row", c:"freezetop", i:"❄"},{l:"Freeze First Column", c:"freezeleft", i:"❄"},{sep:1},{l:"Unfreeze Panes", c:"unfreeze", i:"✕"}],
  catfin:[{l:"PMT (payment)", c:"insertfn", a:"PMT"},{l:"FV (future value)", c:"insertfn", a:"FV"},{l:"NPV (net present value)", c:"insertfn", a:"NPV"}],
  catlog:[{l:"IF", c:"insertfn", a:"IF"},{l:"AND", c:"insertfn", a:"AND"},{l:"OR", c:"insertfn", a:"OR"},{l:"NOT", c:"insertfn", a:"NOT"},{l:"IFERROR", c:"insertfn", a:"IFERROR"}],
  cattext:[{l:"CONCATENATE", c:"insertfn", a:"CONCATENATE"},{l:"LEFT", c:"insertfn", a:"LEFT"},{l:"RIGHT", c:"insertfn", a:"RIGHT"},{l:"MID", c:"insertfn", a:"MID"},{l:"LEN", c:"insertfn", a:"LEN"},{l:"UPPER", c:"insertfn", a:"UPPER"},{l:"LOWER", c:"insertfn", a:"LOWER"},{l:"TRIM", c:"insertfn", a:"TRIM"},{l:"TEXT", c:"insertfn", a:"TEXT"},{l:"VALUE", c:"insertfn", a:"VALUE"}],
  catdate:[{l:"TODAY", c:"insertfn", a:"TODAY"},{l:"NOW", c:"insertfn", a:"NOW"},{l:"DATE", c:"insertfn", a:"DATE"},{l:"YEAR", c:"insertfn", a:"YEAR"},{l:"MONTH", c:"insertfn", a:"MONTH"},{l:"DAY", c:"insertfn", a:"DAY"},{l:"DAYS", c:"insertfn", a:"DAYS"}],
  catlookup:[{l:"VLOOKUP", c:"insertfn", a:"VLOOKUP"},{l:"HLOOKUP", c:"insertfn", a:"HLOOKUP"},{l:"INDEX", c:"insertfn", a:"INDEX"},{l:"MATCH", c:"insertfn", a:"MATCH"},{l:"SUMIF", c:"insertfn", a:"SUMIF"},{l:"COUNTIF", c:"insertfn", a:"COUNTIF"},{l:"AVERAGEIF", c:"insertfn", a:"AVERAGEIF"}],
  catmath:[{l:"SUM", c:"insertfn", a:"SUM"},{l:"AVERAGE", c:"insertfn", a:"AVERAGE"},{l:"ROUND", c:"insertfn", a:"ROUND"},{l:"ABS", c:"insertfn", a:"ABS"},{l:"SQRT", c:"insertfn", a:"SQRT"},{l:"POWER", c:"insertfn", a:"POWER"},{l:"MOD", c:"insertfn", a:"MOD"},{l:"INT", c:"insertfn", a:"INT"}],
  zoommenu:[{l:"Zoom…", c:"zoomdialog"},{l:"100%", c:"zoom100"},{l:"Zoom to Selection", c:"zoomsel"}],
};

/* tab order + groups exactly as Microsoft Excel */
const RIBBON = [
  { name:"Home", groups:[
    { g:"Clipboard", rows:[
      [ {t:"btn", cmd:"paste", icon:I.paste, label:"Paste", menu:"paste", big:1},
        {t:"stack", items:[
          {t:"ico", cmd:"cut", icon:I.cut, title:"Cut (Ctrl+X)"},
          {t:"ico", cmd:"copy", icon:I.copy, title:"Copy (Ctrl+C)"},
          {t:"ico", cmd:"formatpainter", icon:I.brush, title:"Format Painter"},
        ]} ],
    ]},
    { g:"Font", launcher:"fmtdialog", rows:[
      [ {t:"fontsel"}, {t:"sizesel"},
        {t:"ico", cmd:"incfont", icon:I.growf, title:"Increase Font Size"},
        {t:"ico", cmd:"decfont", icon:I.shrinkf, title:"Decrease Font Size"} ],
      [ {t:"ico", cmd:"bold", icon:I.bold, title:"Bold (Ctrl+B)", style:"font-weight:800"},
        {t:"ico", cmd:"italic", icon:I.italic, title:"Italic (Ctrl+I)", style:"font-style:italic"},
        {t:"ico", cmd:"underline", icon:I.under, title:"Underline (Ctrl+U)", style:"text-decoration:underline", menu:"underline"},
        {t:"ico", cmd:"border", icon:I.border, title:"Borders", menu:"borders"},
        {t:"color", cmd:"bcolor", icon:I.bcolor, label:"Fill Color", def:"#ffff00"},
        {t:"color", cmd:"fcolor", icon:I.fcolor, label:"Font Color", def:"#c00000"} ],
    ]},
    { g:"Alignment", launcher:"fmtdialog", rows:[
      [ {t:"ico", cmd:"valigntop", icon:I.at, title:"Top Align"},
        {t:"ico", cmd:"valignmiddle", icon:I.am, title:"Middle Align"},
        {t:"ico", cmd:"valignbottom", icon:I.ab, title:"Bottom Align"},
        {t:"ico", cmd:"fmtdialog", icon:I.orient, title:"Orientation (Format Cells…)"} ],
      [ {t:"ico", cmd:"alignleft", icon:I.al, title:"Align Left"},
        {t:"ico", cmd:"aligncenter", icon:I.ac, title:"Center"},
        {t:"ico", cmd:"alignright", icon:I.ar, title:"Align Right"},
        {t:"ico", cmd:"inddec", icon:I.inddec, title:"Decrease Indent"},
        {t:"ico", cmd:"indinc", icon:I.indinc, title:"Increase Indent"} ],
      [ {t:"btn", cmd:"wrap", icon:I.wrap, label:"Wrap Text", small:1, style:"width:70px"},
        {t:"btn", cmd:"mergecenter", icon:I.merge, label:"Merge & Center", small:1, menu:"merge", style:"width:70px"} ],
    ]},
    { g:"Number", launcher:"fmtdialog", rows:[
      [ {t:"numsel"} ],
      [ {t:"ico", cmd:"currency", icon:I.cur, title:"Accounting Number Format"},
        {t:"ico", cmd:"percent", icon:I.pct, title:"Percent Style (Ctrl+Shift+%)"},
        {t:"ico", cmd:"comma", icon:I.comma, title:"Comma Style"},
        {t:"ico", cmd:"decplus", icon:I.decp, title:"Increase Decimal", cls:"tiny"},
        {t:"ico", cmd:"decminus", icon:I.decm, title:"Decrease Decimal", cls:"tiny"} ],
    ]},
    { g:"Styles", rows:[
      [ {t:"btn", cmd:"cfgt", icon:I.cond, label:"Conditional Formatting", menu:"condfmt", big:1},
        {t:"btn", cmd:"tablestyle", icon:I.table, label:"Format as Table", menu:"astable", big:1},
        {t:"btn", cmd:"cellstyle", icon:I.cstyle, label:"Cell Styles", menu:"cellstyles", big:1} ],
    ]},
    { g:"Cells", rows:[
      [ {t:"btn", cmd:"insertcells", icon:I.ins, label:"Insert", menu:"insertcells", big:1},
        {t:"btn", cmd:"deletecells", icon:I.del, label:"Delete", menu:"deletecells", big:1},
        {t:"btn", cmd:"fmtdialog", icon:I.fmt, label:"Format", menu:"format", big:1} ],
    ]},
    { g:"Editing", rows:[
      [ {t:"btn", cmd:"autosum", icon:I.sum, label:"AutoSum", menu:"autosum", big:1},
        {t:"btn", cmd:"filldown", icon:I.fill, label:"Fill", menu:"fill", big:1},
        {t:"btn", cmd:"clearall", icon:I.clear, label:"Clear", menu:"clear", big:1},
        {t:"btn", cmd:"sortasc", icon:I.sortfilter, label:"Sort & Filter", menu:"sortfilter", big:1},
        {t:"btn", cmd:"finddialog", icon:I.find, label:"Find & Select", menu:"findsel", big:1} ],
    ]},
  ]},
  { name:"Insert", groups:[
    { g:"Tables", items:[
      {t:"btn", cmd:"tablestyle", icon:I.table2, label:"Table", menu:"astable"},
    ]},
    { g:"Charts", items:[
      {t:"btn", cmd:"chartbar", icon:I.chart, label:"Column or Bar Chart"},
      {t:"btn", cmd:"chartline", icon:I.line, label:"Line or Area Chart"},
      {t:"btn", cmd:"chartpie", icon:I.pie, label:"Pie or Doughnut Chart"},
    ]},
    { g:"Symbols", items:[
      {t:"btn", cmd:"symdialog", icon:I.symbol, label:"Symbol"},
    ]},
  ]},
  { name:"Page Layout", groups:[
    { g:"Themes", items:[
      {t:"btn", cmd:"settheme", icon:I.theme, label:"Themes", menu:"themes"},
    ]},
    { g:"Page Setup", launcher:"pagesetup", items:[
      {t:"btn", cmd:"setmargins", icon:I.margins, label:"Margins", menu:"margins"},
      {t:"btn", cmd:"setorient", icon:I.orient, label:"Orientation", menu:"orient"},
      {t:"btn", cmd:"setpsize", icon:I.size, label:"Size", menu:"psize"},
      {t:"btn", cmd:"setprintarea", icon:I.printarea, label:"Print Area", menu:"printarea"},
      {t:"btn", cmd:"setfit", icon:I.breaks, label:"Breaks"},
    ]},
    { g:"Scale to Fit", launcher:"scaledialog", items:[
      {t:"btn", cmd:"setfit", icon:I.scale, label:"Width", menu:"widthfit"},
      {t:"btn", cmd:"zoomdialog", icon:"%", label:"Scale"},
    ]},
    { g:"Sheet Options", items:[
      {t:"chk", cmd:"gridlines", label:"View Gridlines"},
      {t:"chk", cmd:"gridprint", label:"Print Gridlines"},
      {t:"chk", cmd:"headings", label:"View Headings"},
    ]},
  ]},
  { name:"Formulas", groups:[
    { g:"Function Library", launcher:"fxdialog", items:[
      {t:"btn", cmd:"fxdialog", icon:I.fx, label:"Insert Function"},
      {t:"btn", cmd:"autosum", icon:I.sum, label:"AutoSum", menu:"autosum"},
      {t:"btn", cmd:"catfin", icon:I.catfin, label:"Financial", small:1, menu:"catfin"},
      {t:"btn", cmd:"catlog", icon:I.catlog, label:"Logical", small:1, menu:"catlog"},
      {t:"btn", cmd:"cattext", icon:I.cattext, label:"Text", small:1, menu:"cattext"},
      {t:"btn", cmd:"catdate", icon:I.catdate, label:"Date & Time", small:1, menu:"catdate"},
      {t:"btn", cmd:"catlookup", icon:I.catlookup, label:"Lookup & Reference", small:1, menu:"catlookup"},
      {t:"btn", cmd:"catmath", icon:I.catmath, label:"Math & Trig", small:1, menu:"catmath"},
    ]},
    { g:"Formula Auditing", items:[
      {t:"btn", cmd:"showformulas", icon:I.audit, label:"Show Formulas"},
      {t:"btn", cmd:"stats", icon:I.stats, label:"Error Checking"},
    ]},
    { g:"Calculation", launcher:"calcopts", items:[
      {t:"btn", cmd:"recalc", icon:I.recalc, label:"Calculate Now"},
    ]},
  ]},
  { name:"Data", groups:[
    { g:"Sort & Filter", launcher:"customsort", items:[
      {t:"btn", cmd:"sortasc", icon:I.sortfilter, label:"Sort A to Z"},
      {t:"btn", cmd:"sortdesc", icon:I.sortfilter, label:"Sort Z to A"},
      {t:"btn", cmd:"filter", icon:I.filter, label:"Filter"},
      {t:"btn", cmd:"clearfilter", icon:"✕", label:"Clear"},
    ]},
    { g:"Data Tools", items:[
      {t:"btn", cmd:"texttocolumns", icon:I.textcol, label:"Text to Columns"},
      {t:"btn", cmd:"deldup", icon:I.dup, label:"Remove Duplicates"},
      {t:"btn", cmd:"datavalidation", icon:I.valid, label:"Data Validation"},
    ]},
  ]},
  { name:"Review", groups:[
    { g:"Proofing", items:[
      {t:"btn", cmd:"spell", icon:I.spell, label:"Spelling"},
      {t:"btn", cmd:"stats", icon:I.stats, label:"Workbook Statistics"},
    ]},
    { g:"Comments", items:[
      {t:"btn", cmd:"comment", icon:I.comment, label:"New Comment"},
      {t:"btn", cmd:"delcomment", icon:I.delcomment, label:"Delete"},
    ]},
    { g:"Protect", items:[
      {t:"btn", cmd:"protect", icon:I.protect, label:"Protect Sheet"},
    ]},
  ]},
  { name:"View", groups:[
    { g:"Workbook Views", items:[
      {t:"btn", cmd:"viewnormal", icon:I.normal, label:"Normal"},
      {t:"btn", cmd:"viewpagebreak", icon:I.pagebreak, label:"Page Break Preview"},
    ]},
    { g:"Show", items:[
      {t:"chk", cmd:"gridlines", label:"Gridlines"},
      {t:"chk", cmd:"formulabar", label:"Formula Bar"},
      {t:"chk", cmd:"headings", label:"Headings"},
    ]},
    { g:"Zoom", items:[
      {t:"btn", cmd:"zoomdialog", icon:I.zoom, label:"Zoom"},
      {t:"btn", cmd:"zoom100", icon:I.zoom100, label:"100%"},
      {t:"btn", cmd:"zoomsel", icon:I.zoomsel, label:"Zoom to Selection"},
    ]},
    { g:"Window", items:[
      {t:"btn", cmd:"freezedlg", icon:I.freeze, label:"Freeze Panes", menu:"freeze"},
    ]},
  ]},
  { name:"Help", groups:[
    { g:"Help", items:[
      {t:"btn", cmd:"help", icon:I.help, label:"Help"},
      {t:"btn", cmd:"about", icon:I.about, label:"About"},
    ]},
  ]},
];

let activeTab = "Home";
/* buttons whose main area opens their dropdown instead of running a command */
const MENU_ONLY = new Set(["cfgt","tablestyle","cellstyle","insertcells","deletecells","filldown",
  "sortasc","finddialog","clearall","settheme","setmargins","setorient","setpsize","setprintarea",
  "setfit","freezedlg","catfin","catlog","cattext","catdate","catlookup","catmath"]);
function renderTabs(){
  const el = $("#tabs");
  el.innerHTML =
    `<button class="tab file-tab" data-file="1">File</button>` +
    RIBBON.map(t=>`<button class="tab${t.name===activeTab?" active":""}" data-tab="${t.name}">${t.name}</button>`).join("");
  const inp = $("#tellmeinput");
  if(inp){
    inp.oninput = ()=> tellMe(inp.value);
    inp.onkeydown = e=>{ if(e.key==="Escape"){ inp.value=""; hideMenu(); inp.blur(); } };
  }
}
function showTab(name){
  activeTab = name;
  $$("#tabs .tab").forEach(b=>b.classList.toggle("active", b.dataset.tab===name));
  renderRibbon(name);
  $("#ribbon").classList.add("active");
}
function renderRibbon(name){
  const el = $("#ribbon");
  const tab = RIBBON.find(t=>t.name===name) || RIBBON[0];
  el.innerHTML = tab.groups.map(g=>{
    const body = g.rows
      ? `<div class="grows">${g.rows.map(row=>`<div class="grow-row">${row.map(itemHtml).join("")}</div>`).join("")}</div>`
      : `<div class="grow">${(g.items||[]).map(itemHtml).join("")}</div>`;
    return `<div class="rgroup">${body}`
      + (g.launcher?`<button class="launcher" data-cmd="${g.launcher}" title="More">⤡</button>`:"")
      + `<div class="glabel">${g.g}</div></div>`;
  }).join("");
  el.classList.add("active");
  syncToolbar();
  syncChecks();
}
function itemHtml(it){
  const t = it.title || it.label || "";
  if(it.t==="btn"){
    const small = it.small ? " small" : (it.big ? " big" : "");
    const arrow = it.menu ? `<span class="ar" data-menu="${it.menu}">▾</span>` : "";
    const lbl = it.label ? `<span>${it.label}${it.big?arrow:""}</span>` : "";
    return `<button class="rbtn${small}" data-cmd="${it.cmd}"${it.menu?` data-hasmenu="1"`:""} title="${t}" style="${it.style||""}"><span class="ic">${it.icon}</span>${lbl}${it.big?"":arrow}</button>`;
  }
  if(it.t==="ico")
    return `<button class="rbtn ico${it.cls?" "+it.cls:""}" data-cmd="${it.cmd}"${it.menu?` data-hasmenu="1"`:""} title="${t}" style="${it.style||""}"><span class="ic">${it.icon}</span>${it.menu?`<span class="ar" data-menu="${it.menu}">▾</span>`:""}</button>`;
  if(it.t==="stack")
    return `<div class="gstack">${it.items.map(itemHtml).join("")}</div>`;
  if(it.t==="chk")
    return `<label class="rbtn small" title="${t}" style="cursor:pointer"><input type="checkbox" data-chk="${it.cmd}" style="margin:0 3px 0 0">${it.label}</label>`;
  if(it.t==="fontsel")
    return `<select class="rsel w" data-cmd="fontname" title="Font">${FONT_LIST.map(f=>`<option>${f}</option>`).join("")}</select>`;
  if(it.t==="sizesel")
    return `<select class="rsel n" data-cmd="fontsize" title="Font Size">${SIZE_LIST.map(s=>`<option>${s}</option>`).join("")}</select>`;
  if(it.t==="numsel")
    return `<select class="rsel s" data-cmd="numfmt" title="Number Format">${Object.entries(NUMFMTS).map(([k,v])=>`<option value="${k}">${v.label}</option>`).join("")}</select>`;
  if(it.t==="color")
    return `<button class="colorbtn" data-colorcmd="${it.cmd}" data-def="${it.def}" title="${t}"><span class="ic" style="font-size:14px;line-height:15px">${it.icon}</span><span class="bar" style="background:${it.def}"></span></button>`;
  return "";
}

/* ---------- dropdown menus ---------- */
let menuFor = null;
function openMenu(anchor, key){
  const items = MENUS[key];
  if(!items) return;
  const pop = $("#ctxpop");
  pop.innerHTML = items.map(it => it.sep
    ? `<div class="psep"></div>`
    : `<button class="pitem" data-menucmd="${it.c}"${it.a!==undefined?` data-menuarg="${it.a}"`:""}><span class="pi">${it.i||""}</span><span>${it.l}</span></button>`
  ).join("");
  const r = anchor.getBoundingClientRect();
  pop.style.left = Math.min(r.left, innerWidth-250)+"px";
  pop.style.top = (r.bottom+2)+"px";
  pop.classList.add("show");
  menuFor = key;
}
function hideMenu(){ $("#ctxpop").classList.remove("show"); menuFor=null; }
document.addEventListener("click", e=>{
  const mi = e.target.closest("[data-menucmd]");
  if(mi){
    const cmd = mi.dataset.menucmd, arg = mi.dataset.menuarg;
    hideMenu();
    runCmd(cmd, arg);
    return;
  }
  const ar = e.target.closest("[data-menu]");
  if(ar){ e.stopPropagation(); openMenu(ar, ar.dataset.menu); return; }
  if(!e.target.closest("#ctxpop")) hideMenu();
});

/* ---------- color popover ---------- */
let colorCmd = null;
const SWATCHES = ["#ffffff","#f2f2f2","#d9d9d9","#bfbfbf","#a6a6a6","#808080","#595959","#404040","#262626","#000000",
"#e7e6e6","#d0cece","#aeaaaa","#757070","#3a3838","#44546a","#d6dce4","#adb9ca","#8496b0","#323e4f",
"#fff2cc","#ffe699","#ffd966","#ffc000","#bf9000","#7f6000","#fbe5d6","#f8cbad","#f4b183","#ed7d31",
"#fce4d6","#c65911","#e2efda","#c6e0b4","#a9d08e","#70ad47","#548235","#375623","#ddebf7","#bdd7ee",
"#9bc2e6","#5b9bd5","#2e75b6","#1f4e79","#c00000","#ff0000","#ffc000","#92d050","#00b050","#00b0f0"];
function openColorPop(anchor, cmd, def){
  colorCmd = cmd;
  const pop = $("#colorpop");
  pop.innerHTML = `<div style="padding:2px 6px 4px;font-size:11px;color:var(--muted)">Theme Colors</div>
  <div class="swgrid">${SWATCHES.map(s=>`<div class="sw" data-sw="${s}" style="background:${s}"></div>`).join("")}</div>
  <div style="display:flex;gap:4px;align-items:center;padding:6px 6px 2px">
    <input type="color" id="cpick" value="${def||"#000000"}" style="width:34px;height:22px;padding:0;border:1px solid var(--border);background:none">
    <button class="btn primary" id="cok" style="flex:1;padding:3px">Apply</button></div>`;
  const r = anchor.getBoundingClientRect();
  pop.style.left = Math.min(r.left, innerWidth-240)+"px";
  pop.style.top = (r.bottom+2)+"px";
  pop.classList.add("show");
}
document.addEventListener("click", e=>{
  const sw = e.target.closest("[data-sw]");
  if(sw){ applyColor(colorCmd, sw.dataset.sw); $("#colorpop").classList.remove("show"); return; }
  if(e.target.id==="cok"){ applyColor(colorCmd, $("#cpick").value); $("#colorpop").classList.remove("show"); return; }
  const cc = e.target.closest("[data-colorcmd]");
  if(cc){ openColorPop(cc, cc.dataset.colorcmd, cc.dataset.def); return; }
  if(!e.target.closest("#colorpop")) $("#colorpop").classList.remove("show");
});
function applyColor(cmd, color){
  if(cmd==="fcolor") setStyleOnSel({color});
  else if(cmd==="bcolor") setStyleOnSel({bg:color});
}

/* ---------- delegated handlers ---------- */
document.addEventListener("click", e=>{
  const tab = e.target.closest(".tab");
  if(tab){
    if(tab.dataset.file){ openBackstage("Info"); return; }
    showTab(tab.dataset.tab); return;
  }
  const chk = e.target.closest("[data-chk]");
  if(chk) return; // handled by change
  const b = e.target.closest(".rbtn[data-cmd], .launcher[data-cmd]");
  if(b && !e.target.closest("[data-menu]")){
    const cmd = b.dataset.cmd;
    if(MENU_ONLY.has(cmd)){ openMenu(b, b.dataset.menu || cmd); return; }
    runCmd(cmd);
    return;
  }
});
document.addEventListener("change", e=>{
  const chk = e.target.closest("[data-chk]");
  if(chk){ runCmd(chk.dataset.chk); return; }
  const s = e.target.closest("select[data-cmd]");
  if(!s) return;
  runCmd(s.dataset.cmd, s.value);
});

/* ---------- checkbox sync ---------- */
function syncChecks(){
  const map = {
    gridlines: typeof GRID_ON!=="undefined" && GRID_ON,
    gridprint: !!(sheet() && sheet().pageSetup && sheet().pageSetup.gridPrint),
    headings: typeof HEADINGS_ON!=="undefined" && HEADINGS_ON,
    formulabar: typeof FXBAR_ON!=="undefined" && FXBAR_ON,
  };
  $$("#ribbon [data-chk]").forEach(c=>{ c.checked = !!map[c.dataset.chk]; });
}

/* ---------- Tell me / command search ---------- */
function cmdIndex(){
  const out = [];
  RIBBON.forEach(t=>t.groups.forEach(g=>g.items.forEach(it=>{
    if(it.label && it.cmd) out.push({label:it.label, cmd:it.cmd, tab:t.name, group:g.g});
  })));
  Object.keys(MENUS).forEach(k=>MENUS[k].forEach(m=>{
    if(m.l && m.c) out.push({label:m.l, cmd:m.c, arg:m.a, tab:"", group:""});
  }));
  return out;
}
function tellMe(q){
  const pop = $("#ctxpop");
  const ql = (q||"").trim().toLowerCase();
  if(!ql){ hideMenu(); return; }
  const seen = new Set();
  const hits = cmdIndex().filter(x=>{
    if(!x.label.toLowerCase().includes(ql)) return false;
    if(seen.has(x.label)) return false;
    seen.add(x.label); return true;
  }).slice(0,10);
  if(!hits.length){ hideMenu(); return; }
  pop.innerHTML = hits.map(h=>`<button class="pitem" data-menucmd="${h.cmd}"${h.arg!==undefined?` data-menuarg="${h.arg}"`:""}><span class="pi">▸</span><span>${h.label}${h.tab?` <span style="color:var(--muted);font-size:11px">(${h.tab} › ${h.group})</span>`:""}</span></button>`).join("");
  const r = $("#tellme").getBoundingClientRect();
  pop.style.left = Math.min(r.left, innerWidth-280)+"px";
  pop.style.top = (r.bottom+2)+"px";
  pop.classList.add("show");
}

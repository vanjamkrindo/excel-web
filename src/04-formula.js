/* ============ 04 · Mesin Rumus ============ */
let evalCache = {};

function evalCell(r,c){
  const k = ck(r,c);
  if(k in evalCache) return evalCache[k];
  return rawCellValue(getCell(r,c));
}
function rawCellValue(cell){
  if(!cell) return "";
  if(cell.f) return "#CIRC?"; // replaced during recalcAll
  return cell.v!==undefined ? cell.v : "";
}

function recalcAll(){
  evalCache = {};
  /* drop the side-by-side caches too, so both panes recompute after an edit */
  if(typeof WORKBOOKS !== "undefined") WORKBOOKS.forEach(w=>{ w._sc = null; });
  const sh = sheet();
  const visiting = new Set();
  function ev(r,c){
    const k = ck(r,c);
    if(k in evalCache) return evalCache[k];
    const cell = sh.cells[k];
    if(!cell) return "";
    if(cell.f){
      if(visiting.has(k)) return "#CIRC!";
      visiting.add(k);
      let v;
      try { v = evalFormula(cell.f, r, c); }
      catch(err){ v = "#ERROR!"; }
      visiting.delete(k);
      if(typeof v==="string" && v.startsWith("#")) { /* keep */ }
      evalCache[k] = v;
      return v;
    }
    evalCache[k] = cell.v!==undefined ? cell.v : "";
    return evalCache[k];
  }
  wb._eval = ev; // expose for io
  for(const k in sh.cells){ const [r,c]=k.split(",").map(Number); ev(r,c); }
}

/* Evaluate one cell of a specific workbook/sheet, independent of what is active.
   Used by the side-by-side view so both panes show computed values. */
function computeCell(w, r, c){
  if(!w) return "";
  const sh = w.sheets[Math.min((w.cur && w.cur.sheet) || 0, w.sheets.length-1)];
  if(!sh) return "";
  const cell = sh.cells[ck(r,c)];
  if(!cell) return "";
  if(!cell.f) return cell.v!==undefined ? cell.v : "";
  /* The cached value is keyed per workbook; recalcAll clears it (see below). */
  if(!w._sc) w._sc = {};
  const key = ck(r,c);
  if(key in w._sc) return w._sc[key];
  const prev = {wb: wb, cur: cur};
  let v;
  try{
    wb = w; cur = {sheet: w.cur ? w.cur.sheet : 0, r:r, c:c};
    v = evalFormula(cell.f, r, c);
  }catch(e){ v = "#ERROR!"; }
  finally{ wb = prev.wb; cur = prev.cur; }
  w._sc[key] = v;
  return v;
}

/* --- formula parser --- */
function evalFormula(f, r0, c0){
  let src = f.replace(/^=/,"");
  src = src.replace(/(\$?[A-Z]{1,3}\$?\d+):(\$?[A-Z]{1,3}\$?\d+)/gi,(m,a,b)=>{
    const p1=parseRef(a), p2=parseRef(b);
    if(!p1||!p2) return m;
    return `RTV(${JSON.stringify(rngVals(p1,p2))},${p2.c-p1.c+1},${p2.r-p1.r+1})`;
  });
  src = src.replace(/(?<![A-Za-z0-9$!])(\$?[A-Z]{1,3}\$?\d+)(?![\w(])/gi, (m,ref)=>{
    const p = parseRef(ref);
    if(!p) return m;
    return `C(${p.r},${p.c})`;
  });
  // pemisah argumen gaya Indonesia (;) -> koma, di luar tanda kutip
  src = src.replace(/"[^"]*"|;/g, m => m===";" ? "," : m);
  // nama fungsi -> FN.NAMA
  const fnRe = new RegExp("(?<![\\w.])([A-Z][A-Z0-9.]*)\\s*\\(","g");
  src = src.replace(fnRe, (m,name)=> FUNCS[name.toUpperCase()] ? "FN."+name.toUpperCase()+"(" : m);
  const C = (r,c)=>{
    if(r<0||c<0||r>=NROWS||c>=NCOLS) return "#REF!";
    const cell = sheet().cells[ck(r,c)];
    if(!cell) return "";
    if(cell.f){
      if(ck(r,c) in evalCache) return evalCache[ck(r,c)];
      let v; try{ v = evalFormula(cell.f, r, c);}catch(e){ v="#VALUE!"; }
      evalCache[ck(r,c)] = v;
      return v;
    }
    return cell.v!=null?cell.v:"";
  };
  const RTV = (vals,w,h)=>{ rangeDims={w,h}; return vals.map(x=> typeof x==="string" && x!=="" && !isNaN(parseFloat(x)) ? parseFloat(x) : x); };
  const fn = evalFnCache.get(src) || (()=>{
    const f = new Function("C","RTV","FN","PI","return ("+src+");");
    evalFnCache.set(src, f); return f;
  })();
  return fn(C, RTV, FUNCS, Math.PI);
}
const evalFnCache = new Map();
function rngVals(p1,p2){
  const out=[];
  for(let r=Math.min(p1.r,p2.r);r<=Math.max(p1.r,p2.r);r++)
    for(let c=Math.min(p1.c,p2.c);c<=Math.max(p1.c,p2.c);c++){
      const cell = sheet().cells[ck(r,c)];
      let v;
      if(!cell) v = "";
      else if(cell.f){
        if(ck(r,c) in evalCache) v = evalCache[ck(r,c)];
        else { try{ v = evalFormula(cell.f,r,c);}catch(e){ v="#VALUE!"; } }
      }
      else v = cell.v!=null?cell.v:"";
      out.push(typeof v==="string" && v!=="" && !isNaN(parseFloat(v)) ? parseFloat(v) : v);
    }
  return out;
}

const FUNCS = {
  SUM:(...a)=>numFlat(a).reduce((x,y)=>x+y,0),
  AVERAGE:(...a)=>{const n=numFlat(a); return n.length? n.reduce((x,y)=>x+y,0)/n.length : "#DIV/0!";},
  MIN:(...a)=>{const n=numFlat(a); return n.length?Math.min(...n):0;},
  MAX:(...a)=>{const n=numFlat(a); return n.length?Math.max(...n):0;},
  COUNT:(...a)=>numFlat(a).length,
  COUNTA:(...a)=>flat(a).filter(x=>x!==""&&x!=null).length,
  COUNTBLANK:(a)=>flat([a]).filter(x=>x===""||x==null).length,
  ABS:x=>Math.abs(x), SQRT:x=>Math.sqrt(x),
  POWER:(x,y)=>Math.pow(x,y), MOD:(x,y)=>x%y, INT:x=>Math.floor(x),
  ROUND:(x,n=0)=>{const p=Math.pow(10,n); return Math.round(x*p)/p;},
  ROUNDUP:(x,n=0)=>{const p=Math.pow(10,n); return Math.ceil(Math.abs(x)*p)/p*Math.sign(x);},
  ROUNDDOWN:(x,n=0)=>{const p=Math.pow(10,n); return Math.floor(Math.abs(x)*p)/p*Math.sign(x);},
  CEILING:(x,s=1)=>Math.ceil(x/s)*s, FLOOR:(x,s=1)=>Math.floor(x/s)*s,
  SQRTPI:x=>Math.sqrt(x*Math.PI),
  EXP:Math.exp, LN:x=>Math.log(x), LOG10:Math.log10, LOG10X:null,
  SIN:Math.sin, COS:Math.cos, TAN:Math.tan, ATAN:Math.atan, ASIN:Math.asin, ACOS:Math.acos,
  IF:(c,a,b="")=> c? a : b,
  IFERROR:(v,alt)=> (typeof v==="string"&&v.startsWith("#")) ? alt : v,
  AND:(...a)=>flat(a).every(Boolean), OR:(...a)=>flat(a).some(Boolean), NOT:x=>!x,
  TRUE:()=>true, FALSE:()=>false,
  CONCAT:(...a)=>flat(a).join(""), CONCATENATE:(...a)=>flat(a).join(""),
  LEFT:(s,n=1)=>String(s).slice(0,n), RIGHT:(s,n=1)=>String(s).slice(-n),
  MID:(s,i,n)=>String(s).slice(i-1,i-1+n),
  LEN:s=>String(s).length, UPPER:s=>String(s).toUpperCase(), LOWER:s=>String(s).toLowerCase(),
  PROPER:s=>String(s).replace(/\w\S*/g,w=>w[0].toUpperCase()+w.slice(1).toLowerCase()),
  TRIM:s=>String(s).trim().replace(/\s+/g," "),
  SUBSTITUTE:(s,a,b)=>String(s).split(a).join(b),
  FIND:(a,s,i=1)=>String(s).indexOf(a,i-1)+1,
  SEARCH:(a,s,i=1)=>String(s).toLowerCase().indexOf(String(a).toLowerCase(),i-1)+1,
  REPT:(s,n)=>String(s).repeat(n),
  TEXT:(v,f)=>{ // minimal: number with separators / percent
    if(f.includes("%")) return (v*100).toLocaleString("id-ID")+"%";
    return typeof v==="number"? v.toLocaleString("id-ID") : String(v);
  },
  VALUE:s=>parseFloat(s),
  TODAY:()=>toSerial(new Date(new Date().toDateString())),
  NOW:()=>toSerial(new Date()),
  YEAR:d=>serialToDate(d).getFullYear(), MONTH:d=>serialToDate(d).getMonth()+1, DAY:d=>serialToDate(d).getDate(),
  DATE:(y,m,d)=>toSerial(new Date(y,m-1,d)),
  WEEKDAY:(d)=>serialToDate(d).getDay()+1,
  RAND:()=>Math.random(),
  RANDBETWEEN:(a,b)=>Math.floor(Math.random()*(b-a+1))+a,
  VLOOKUP:(k,tab,ci)=>vlookupImpl(k,tab,ci),
  HLOOKUP:(k,tab,ri)=>hlookupImpl(k,tab,ri),
  INDEX:(tab,r,c)=>indexImpl(tab,r,c),
  MATCH:(k,vec)=>matchImpl(k,vec),
  SUMIF:(rng,crit,sumrng)=>sumifImpl(rng,crit,sumrng),
  COUNTIF:(rng,crit)=>countifImpl(rng,crit),
  AVERAGEIF:(rng,crit)=>avgifImpl(rng,crit),
  MEDIAN:(...a)=>{const n=numFlat(a).sort((x,y)=>x-y); if(!n.length)return 0; const m=n.length>>1; return n.length%2?n[m]:(n[m-1]+n[m])/2;},
  STDEV:(...a)=>{const n=numFlat(a); if(n.length<2)return 0; const mu=n.reduce((x,y)=>x+y,0)/n.length; return Math.sqrt(n.reduce((s,x)=>s+(x-mu)**2,0)/(n.length-1));},
  VAR:(...a)=>{const n=numFlat(a); if(n.length<2)return 0; const mu=n.reduce((x,y)=>x+y,0)/n.length; return n.reduce((s,x)=>s+(x-mu)**2,0)/(n.length-1);},
  PRODUCT:(...a)=>numFlat(a).reduce((x,y)=>x*y,1),
  PMT:(r,n,pv)=>{ if(!r) return pv/n; return r*pv/(1-Math.pow(1+r,-n)); },
  FV:(r,n,pmt,pv=0)=>{ if(!r) return -(pmt*n+pv); return -(pv*Math.pow(1+r,n)+pmt*((Math.pow(1+r,n)-1)/r)); },
  ISNUMBER:v=>typeof v==="number", ISTEXT:v=>typeof v==="string",
  ISBLANK:v=>v===""||v==null, ISERROR:v=>typeof v==="string"&&v.startsWith("#"),
  NA:()=>"#N/A",
};
function numFlat(a){ return flat(a).filter(x=>typeof x==="number"&&!Number.isNaN(x)); }
function flat(a){ const out=[]; for(const x of a){ if(Array.isArray(x)) out.push(...flat(x)); else out.push(x);} return out; }
function matchCrit(crit){
  if(typeof crit==="string"){
    const m=/^(<=|>=|<>|=|<|>)(.*)$/.exec(crit);
    if(m){ const op=m[1], v=isNaN(parseFloat(m[2]))?m[2]:parseFloat(m[2]);
      return x=>({ "<":x<v, ">":x>v, "<=":x<=v, ">=":x>=v, "<>":x!=v, "=":x==v }[typeof x==="number"&&typeof v==="number" ? op : op==="="?String(x)===String(m[2]):op==="<>"?String(x)!==String(m[2]):false] ?? false);
    }
    const cn=parseFloat(crit);
    if(!isNaN(cn)) return x=>x==cn;
    return x=>String(x).toLowerCase().includes(crit.toLowerCase());
  }
  return x=>x==crit;
}
function sumifImpl(rv,crit,sv){
  const f=matchCrit(crit); let s=0;
  rv.forEach((v,i)=>{ if(f(v)) { const x=sv? sv[i]:v; if(typeof x==="number") s+=x; }});
  return s;
}
function countifImpl(rv,crit){ const f=matchCrit(crit); return rv.filter(f).length; }
function avgifImpl(rv,crit){ const f=matchCrit(crit); const vs=rv.filter(f).filter(x=>typeof x==="number"); return vs.length?vs.reduce((a,b)=>a+b,0)/vs.length:"#DIV/0!"; }
function vlookupImpl(k,tab,ci){
  const w = rangeDims ? rangeDims.w : 1;
  for(let i=0;i<tab.length;i+=w){
    if(String(tab[i]).toLowerCase()===String(k).toLowerCase()) return tab[i+(ci-1)] ?? "#N/A";
  }
  return "#N/A";
}
function hlookupImpl(k,tab,ri){
  const w = rangeDims ? rangeDims.w : 1;
  const hdr = tab.slice(0,w); const idx = hdr.findIndex(h=>String(h).toLowerCase()===String(k).toLowerCase());
  if(idx<0) return "#N/A";
  return tab[(ri-1)*w + idx] ?? "#N/A";
}
function indexImpl(arr,r,c=1){
  if(rangeDims){ const {w}=rangeDims;
    if(r===0) return arr.filter((_,i)=>i%w===(c-1));
    if(c===0) return arr.slice((r-1)*w, (r-1)*w+w);
    return arr[(r-1)*w + (c-1)] ?? "#REF!";
  }
  return arr[r-1] ?? "#N/A";
}
function matchImpl(k,arr){
  const f=matchCrit(k);
  for(let i=0;i<arr.length;i++) if(f(arr[i])) return i+1;
  return "#N/A";
}
let rangeDims = null;

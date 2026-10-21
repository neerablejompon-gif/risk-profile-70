/* ===== Excel parser: แบบฟอร์ม "ผลความเสี่ยง ไตรมาส" (1 ชีท = 1 สาเหตุ, ชื่อชีท "S1 (R1)") ===== */
function parseRiskWorkbook(XLSX, wb){
  const clean = v => v == null ? '' : String(v).replace(/\r/g,'').replace(/[ \t ]+/g,' ').replace(/ *\n */g,'\n').trim();
  const flat = v => clean(v).replace(/\n+/g,' ').trim();
  const stripLabel = (v, re) => flat(v).replace(re,'').trim();
  const num = v => { const n = parseInt(String(v).trim(),10); return n>=1 && n<=5 ? n : null; };
  const warnings = [];

  function grid(ws){
    const ref = ws['!ref']; if(!ref) return {rows:0,get:()=>''};
    const rg = XLSX.utils.decode_range(ref);
    const get = (r,c) => { const cell = ws[XLSX.utils.encode_cell({r:r-1,c})]; if(!cell) return ''; return cell.f && cell.v == null ? '' : cell.v; };
    return {rows: rg.e.r+1, get};
  }
  const A=0,B=1,C=2,D=3,E=4,F=5,G=6,H=7;

  function parseTable(g, startRow){
    // startRow = แถวหัวตาราง "ที่"; อ่านแถวที่คอลัมน์ A เป็นตัวเลข หรือแถวต่อเนื่องที่มีข้อมูลใน B
    const rows=[];
    for(let r=startRow+1; r<=g.rows; r++){
      const a = flat(g.get(r,A));
      if(a && !/^\d+$/.test(a)) break;
      const row = {plan:clean(g.get(r,B)), period:clean(g.get(r,C)), result:clean(g.get(r,D)), issue:clean(g.get(r,E)), solution:clean(g.get(r,F))};
      if(!a && !row.plan && !row.result) { if(rows.length) break; else continue; }
      rows.push(row);
    }
    return rows;
  }

  function parseSheet(ws, name){
    const g = grid(ws);
    const out = {factorCode:'',factorName:'',period:'',strategy:'',tactic:'',cause:{code:'',name:'',owner:'',ra:'',rt:'',likelihoodTitle:'',impactTitle:'',likelihoodActual:'',impactActual:'',levels:[],before:{l:null,i:null},after:{l:null,i:null},plans:[]},weaknesses:[],controls:[]};
    let seenL = 0, inWeak = false;
    for(let r=1; r<=g.rows; r++){
      const a = flat(g.get(r,A)), b = g.get(r,B);
      if(a.startsWith('ปัจจัยเสี่ยงระดับองค์กร') && !out.factorName){
        const m = flat(b).match(/^([A-Za-z]+\d*)\s+(.*)$/); if(m){ out.factorCode=m[1]; out.factorName=m[2]; } else out.factorName=flat(b);
      }
      else if(a.startsWith('ผลการดำเนินงาน') && a.length<20 && !out.period) out.period = flat(b);
      else if(a.startsWith('ยุทธศาสตร์')) out.strategy = flat(b);
      else if(a.startsWith('กลยุทธ์')) out.tactic = flat(b);
      else if(a.startsWith('สาเหตุความเสี่ยง') && !out.cause.name){
        const m = flat(b).match(/^(R\d+)\s*(.*)$/); if(m){ out.cause.code=m[1]; out.cause.name=m[2]; } else out.cause.name=flat(b);
      }
      else if(a.startsWith('เจ้าของปัจจัยเสี่ยง')) out.cause.owner = flat(b);
      if(/^Risk Appetite/i.test(flat(b))) out.cause.ra = stripLabel(b,/^Risk Appetite\s*:\s*/i);
      if(/^Risk Tolerance/i.test(flat(b))) out.cause.rt = stripLabel(b,/^Risk Tolerance\s*:\s*/i);
      if(a === 'ระดับ'){
        out.cause.likelihoodTitle = stripLabel(g.get(r,B),/^โอกาสเกิด\s*:\s*/);
        out.cause.impactTitle = stripLabel(g.get(r,D),/^ผลกระทบ\s*:\s*/);
      }
      if(/^[1-5]$/.test(a) && out.cause.likelihoodTitle && out.cause.levels.length<5 && !out.cause.plans.length){
        out.cause.levels.push({level:+a, likelihood:clean(g.get(r,B)), impact:clean(g.get(r,D))});
        if(!out.cause.likelihoodActual) out.cause.likelihoodActual = clean(g.get(r,C));
        if(!out.cause.impactActual) out.cause.impactActual = clean(g.get(r,E));
      }
      if(flat(g.get(r,F)) === 'โอกาสเกิด'){
        const t = seenL++ === 0 ? out.cause.before : out.cause.after;
        t.l = num(g.get(r+1,F)); t.i = num(g.get(r+1,G));
      }
      if(a.startsWith('จุดอ่อนของการควบคุม')) inWeak = true;
      else if(a) inWeak = false;
      if(inWeak && flat(b)) out.weaknesses.push(flat(b).replace(/^\d+\.\s*/,''));
      if(a.startsWith('ผลการดำเนินงานตามมาตรการปรับปรุง')){ const h = flat(g.get(r+1,A))==='ที่' ? r+1 : r; out.controls = parseTable(g,h).map(x=>({measure:x.plan,period:x.period,result:x.result,issue:x.issue,solution:x.solution})); }
      if(a.startsWith('ผลการดำเนินงานตามแผนจัดการความเสี่ยง')){ const h = flat(g.get(r+1,A))==='ที่' ? r+1 : r; out.cause.plans = parseTable(g,h); }
    }
    if(!out.cause.before.l || !out.cause.before.i) warnings.push(`${name}: ไม่พบค่าระดับความเสี่ยงก่อนบริหาร`);
    if(!out.cause.after.l || !out.cause.after.i) warnings.push(`${name}: ไม่พบค่าระดับความเสี่ยงหลังบริหาร`);
    return out;
  }

  function parseKRI(ws){
    const g = grid(ws), list=[]; let cur=null;
    for(let r=3; r<=g.rows; r++){
      const code = flat(g.get(r,A));
      if(code){ cur = {code, name:flat(g.get(r,B)), strategy:flat(g.get(r,C)), tactic:clean(g.get(r,D)), kris:[]}; list.push(cur); }
      const k = flat(g.get(r,E));
      if(cur && k) cur.kris.push({name:k, ra:flat(g.get(r,F)), rt:flat(g.get(r,G)), owner:flat(g.get(r,H))});
    }
    return list;
  }

  const factors = [], byCode = {};
  let period = '';
  let kriList = [];
  wb.SheetNames.forEach(name=>{
    const ws = wb.Sheets[name];
    if(/^KRI$/i.test(name.trim())){ kriList = parseKRI(ws); return; }
    const m = name.match(/^\s*([A-Za-z]+\d*)\s*\((R\d+)\)\s*$/);
    if(!m){ warnings.push(`ข้ามชีท "${name}" (ชื่อชีทไม่ตรงรูปแบบ "S1 (R1)")`); return; }
    const s = parseSheet(ws, name);
    const code = s.factorCode || m[1];
    if(s.factorCode && !s.factorCode.startsWith(m[1])) warnings.push(`ชีท "${name}" ระบุปัจจัยเสี่ยง ${s.factorCode} ในเนื้อหา จึงจัดไว้ใต้ ${s.factorCode}`);
    if(!period && s.period) period = s.period;
    let f = byCode[code];
    if(!f){ f = byCode[code] = {id:uid(), code, name:s.factorName, strategy:s.strategy, tactic:s.tactic, kris:[], weaknesses:[], controls:[], causes:[]}; factors.push(f); }
    if(!f.weaknesses.length && s.weaknesses.length) f.weaknesses = s.weaknesses;
    if(!f.controls.length && s.controls.length) f.controls = s.controls;
    if(!s.cause.code) s.cause.code = m[2];
    f.causes.push(Object.assign({id:uid()}, s.cause));
  });
  // ผูก KRI (ชีท KRI ใช้รหัส IT1 ส่วนชื่อชีทใช้ IT)
  kriList.forEach(k=>{
    const f = byCode[k.code] || factors.find(x=>k.code.startsWith(x.code) || x.code.startsWith(k.code));
    if(f){ f.kris = k.kris; if(k.code.length>f.code.length) f.code = k.code; if(!f.name) f.name = k.name; }
    else warnings.push(`KRI ${k.code} ไม่มีชีทผลการดำเนินงาน`);
  });
  const pm = period.match(/ไตรมาส\s*(\d)\s*\/\s*(\d{4})/), am = period.match(/ณ\s*เดือน\s*(\S+)\s*(\d{4})/);
  return {factors, warnings, meta:{period, quarter: pm?+pm[1]:null, year: pm?+pm[2]:null, asOfMonth: am?am[1]:'', asOfYear: am?+am[2]:null}};
}

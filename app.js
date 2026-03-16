
import * as pdfjsLib from 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.4.168/build/pdf.min.mjs';
pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdn.jsdelivr.net/npm/pdfjs-dist@4.4.168/build/pdf.worker.min.mjs';

const STORAGE_KEY='ocr_asn_goodmark_final_state';
const DEFAULT_SINGLE=[
  {item:'551727001',rev:'01',packing:280,note:'Default'},
  {item:'524451003',rev:'01',packing:270,note:'Default'},
  {item:'317015001',rev:'01',packing:253,note:'Default'}
];
const DEFAULT_PAIR=[
  {itemA:'545861001',revA:'03',itemB:'545862001',revB:'02',packing:24,note:'XC2 pair'},
  {itemA:'545862001',revA:'02',itemB:'545863001',revB:'01',packing:24,note:'XC2 pair'}
];

const $=id=>document.getElementById(id);
const state=loadState();
let currentFiles=[];
let currentParsed={asnNo:'',rawLines:[],results:[],meta:{},warnings:[]};
let currentPageParses=[];

function loadState(){
  const raw=localStorage.getItem(STORAGE_KEY);
  const s = raw ? JSON.parse(raw) : {singlePacking:[...DEFAULT_SINGLE],pairPacking:[...DEFAULT_PAIR],history:[]};
  return normalizePackingState(s);
}
function saveState(){ normalizePackingState(state); localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); }

function normalizePackingState(s){
  s.singlePacking = (s.singlePacking||[]).map(p=>({
    item: normalizeItem(p.item),
    rev: normalizeRev(p.rev),
    packing: Number(p.packing||0),
    note: p.note||''
  })).filter(p=>/^[35]\d{8}$/.test(p.item) && p.rev && p.packing>0);
  const dedupSingle = new Map();
  s.singlePacking.forEach(p=>dedupSingle.set(`${p.item}|${p.rev}`, p));
  s.singlePacking = Array.from(dedupSingle.values());

  s.pairPacking = (s.pairPacking||[]).map(p=>({
    itemA: normalizeItem(p.itemA),
    revA: normalizeRev(p.revA),
    itemB: normalizeItem(p.itemB),
    revB: normalizeRev(p.revB),
    packing: Number(p.packing||0),
    note: p.note||''
  })).filter(p=>p.itemA && p.revA && p.itemB && p.revB && p.packing>0);
  const dedupPair = new Map();
  s.pairPacking.forEach(p=>dedupPair.set(`${p.itemA}|${p.revA}|${p.itemB}|${p.revB}`, p));
  s.pairPacking = Array.from(dedupPair.values());

  s.history = s.history || [];
  s.historyBatches = s.historyBatches || [];
  return s;
}

function normalizeItem(v){
  let x=String(v||'').toUpperCase().replace(/[^0-9A-Z]/g,'');
  x=x.replace(/O/g,'0').replace(/[IL]/g,'1').replace(/S/g,'5').replace(/B/g,'8');
  return x;
}
function normalizeRev(v){
  let s=String(v||'').toUpperCase().trim();
  s=s.replace(/\bOR\b/g,'01').replace(/\bO1\b/g,'01').replace(/\bOL\b/g,'01');
  const d=s.replace(/\D/g,'');
  return d ? d.padStart(2,'0').slice(-2) : '';
}
function normalizeLineNo(v){
  let s=String(v||'').toUpperCase().replace(/[\s_–—]+/g,'-').replace(/--+/g,'-').trim();
  s=s.replace(/\bGP-?JOB\b/g,'GP JOB').replace(/\bD2-?JOB\b/g,'D2 JOB');
  return s.replace(/[^A-Z0-9 -]/g,'');
}

function calcOrderTotalCartons(rows){
  rows = rows || [];
  const full = rows.reduce((a,r)=>a + Number(r.carton || 0), 0);
  const loose = rows.some(r => Number(r.pcs || 0) > 0) ? 1 : 0;
  return full + loose;
}
function inferMime(name){ return name.toLowerCase().endsWith('.pdf') ? 'application/pdf' : 'image/*'; }

function fixOcrCommonErrors(text){
  text = String(text || '');
  // common OCR fixes in ASN documents
  text = text.replace(/%/g, '6');
  text = text.replace(/\bO(?=\d)/g, '0');
  text = text.replace(/(?<=\d)O\b/g, '0');
  return text;
}

function parseSchedule(line){
  const eta=line.match(/ETA[:\s]+(\d{4}[-/.]\d{2}[-/.]\d{2})\s+([0-2]?\d:\d{2})/i);
  if(eta) return {date:eta[1].replace(/[/.]/g,'-'), times:[eta[2]]};
  return {date:'',times:[]};
}
function itemPatternOk(item, rev){ return /^[35]\d{8}$/.test(item) && /^\d{2}$/.test(rev); }

function setPackingTab(mode){
  const single=mode==='single';
  $('singlePanel').classList.toggle('hidden', !single);
  $('pairPanel').classList.toggle('hidden', single);
  $('showSingleBtn').classList.toggle('ghost', !single);
  $('showPairBtn').classList.toggle('ghost', single);
}

function handleFiles(e){
  const files=Array.from(e.target.files||[]);
  if(currentFiles.length+files.length>50) return alert('Tối đa 50 file mỗi lần scan.');
  files.forEach(file=>currentFiles.push({id:crypto.randomUUID(),name:file.name,type:file.type||inferMime(file.name),file}));
  renderScanPreview();
  e.target.value='';
}
function renderScanPreview(){
  const box=$('scanPreview');
  box.innerHTML='';
  currentFiles.forEach(entry=>{
    const div=document.createElement('div');
    div.className='preview-item';
    if(entry.type.startsWith('image/')){
      const img=document.createElement('img');
      img.src=URL.createObjectURL(entry.file);
      div.appendChild(img);
    }else{
      const ph=document.createElement('div');
      ph.style.cssText='height:120px;display:grid;place-items:center;background:#e5e7eb;border-radius:8px';
      ph.textContent='PDF';
      div.appendChild(ph);
    }
    const n=document.createElement('div');
    n.className='name';
    n.textContent=entry.name;
    const del=document.createElement('button');
    del.className='small danger';
    del.textContent='Xóa';
    del.onclick=()=>{ currentFiles=currentFiles.filter(f=>f.id!==entry.id); renderScanPreview(); };
    div.append(n,del);
    box.appendChild(div);
  });
}
function clearScanState(){
  currentFiles=[]; currentPageParses=[];
  currentParsed={asnNo:'',rawLines:[],results:[],meta:{},warnings:[]};
  $('scanPreview').innerHTML=''; $('qrText').value=''; $('ocrText').value='';
  renderResults();
}
function loadSampleText(){
  $('ocrText').value=`ASN No: CR0372179
ETA: 2026-03-16 06:00
1 63525313-70 545861001 03 366 PC SO: 0062386 C2-001D
2 63528646-55 545861001 03 6 PC SO: 0062386 C2-001D
3 63525313-71 545862001 02 366 PC SO: 0062386 C2-001D
4 63528646-56 545862001 02 6 PC SO: 0062386 C2-001D
5 63525313-72 545863001 01 360 PC SO: 0062386 C2-001D
6 63528646-57 545863001 01 12 PC SO: 0062386 C2-001D`;
  $('qrText').value='CR0372179';
}

async function scanQrFromImages(){
  const imageFiles=currentFiles.filter(f=>f.type.startsWith('image/'));
  if(!imageFiles.length) return alert('Chưa có ảnh để quét QR.');
  let found=[];
  if('BarcodeDetector' in window){
    try{
      const detector=new BarcodeDetector({formats:['qr_code']});
      for(const entry of imageFiles){
        const bmp=await createImageBitmap(entry.file);
        const codes=await detector.detect(bmp);
        codes.forEach(c=>c.rawValue&&found.push(c.rawValue));
      }
    }catch(e){ console.error(e); }
  }
  if(!found.length){
    found=[...new Set((($('ocrText').value||'').match(/\b((CR|CH)\d{7,})\b/gi)||[]).map(s=>s.toUpperCase()))];
  }
  $('qrText').value=found.length?found.join(', '):'Không phát hiện QR trực tiếp.';
}

async function renderPdfToImages(file){
  const arrayBuffer=await file.arrayBuffer();
  const pdf=await pdfjsLib.getDocument({data:arrayBuffer}).promise;
  const blobs=[];
  for(let i=1;i<=pdf.numPages;i++){
    const page=await pdf.getPage(i);
    const viewport=page.getViewport({scale:2});
    const canvas=document.createElement('canvas');
    const ctx=canvas.getContext('2d');
    canvas.width=viewport.width;
    canvas.height=viewport.height;
    await page.render({canvasContext:ctx,viewport}).promise;
    const blob=await new Promise(res=>canvas.toBlob(res,'image/png'));
    blobs.push({blob,pageNo:i});
  }
  return blobs;
}

function parseASNPageText(text, fallbackAsn=''){
  text = fixOcrCommonErrors(text);
  const cleaned=text.replace(/\r/g,'');
  const warnings=[]; const rawLines=[];
  let currentASN=fallbackAsn||''; let currentDate=''; let currentTimes=[];
  let groupCode='';

  cleaned.split('\n').map(s=>s.trim()).filter(Boolean).forEach(line=>{
    const asn=line.match(/\b((CR|CH)\d{7,})\b/i);
    if(asn){ currentASN=asn[1].toUpperCase(); }

    const grp=line.toUpperCase().match(/\b(XC\d-TC\d)\b/);
    if(grp) groupCode=grp[1];

    const sch=parseSchedule(line);
    if(sch.date){ currentDate=sch.date; currentTimes=sch.times; }

    if(/CHECKED|POSTED|SECURITY|CONFIRMED|RECEIVED|TOTAL QUANTITY|DELIVERY NOTE|GOOD MARK|ISSUED BY/i.test(line.toUpperCase())) return;
    if(!currentASN) return;

    let norm=' '+line.toUpperCase()+' ';
    norm=norm.replace(/[|]/g,' ');
    norm=norm.replace(/\bOR\b/g,'01');
    norm=norm.replace(/\bO1\b/g,'01');
    norm=norm.replace(/\bOL\b/g,'01');
    norm=norm.replace(/\bOO\b/g,'00');
    norm=norm.replace(/\s+/g,' ').trim();

    const tokens=norm.split(' ');
    if(tokens.length<5) return;

    let picked=null;
    for(let i=0;i<tokens.length-3;i++){
      const item=normalizeItem(tokens[i]);
      const rev=normalizeRev(tokens[i+1]);
      const qtyTok=tokens[i+2];
      const unitTok=(tokens[i+3]||'').toUpperCase();
      const prev=(tokens[i-1]||'').toUpperCase();

      if(!itemPatternOk(item,rev)) continue;
      if(!/^\d+$/.test(qtyTok)) continue;
      if(unitTok!=='PC') continue;
      if(/PO|SO|ORDER|DOC|NO|XC\d/.test(prev)) continue;

      picked={item,rev,qty:Number(qtyTok),idx:i};
      break;
    }
    if(!picked) return;

    let lineNo='';
    for(const p of [
      /\bC1[\s\-_–—]*\d{3,4}D\b/i,
      /\bC1[\s\-_–—]*\d{3,4}\b/i,
      /\bC2[\s\-_–—]*\d{2,4}D\b/i,
      /\bC2[\s\-_–—]*\d{2,4}\b/i,
      /\bGP[\s\-_–—]*JOB\b/i,
      /\bD2[\s\-_–—]*JOB\b/i
    ]){
      const m=norm.match(p);
      if(m){ lineNo=normalizeLineNo(m[0]); break; }
    }

    rawLines.push({
      asnNo: currentASN,
      itemNo: picked.item,
      rev: picked.rev,
      quantity: picked.qty,
      lineNo,
      groupCode,
      scheduleDate: currentDate,
      scheduleTimes: [...currentTimes]
    });
  });

  return {asnNo:currentASN, groupCode, rawLines, warnings};
}

async function runOcrOnFiles(){
  if(!currentFiles.length) return alert('Chưa chọn file.');
  $('ocrText').value='Đang OCR...';
  currentPageParses=[];
  let combined='';

  for(const entry of currentFiles){
    const fallbackAsn=(entry.name.match(/((CR|CH)\d{7,})/i)||[])[1]?.toUpperCase() || '';
    if(entry.type.startsWith('image/')){
      const result=await Tesseract.recognize(entry.file,'eng',{});
      const text=fixOcrCommonErrors(result.data.text||'');
      combined += `\n===== ${entry.name} =====\n` + text + '\n';
      const parsed=parseASNPageText(text, fallbackAsn);
      currentPageParses.push({source:entry.name, asnNo:parsed.asnNo||fallbackAsn, rawLines:(parsed.rawLines||[]).map(r=>({...r, sourceName: entry.name}))});
    }else if(entry.type.includes('pdf')){
      const pages=await renderPdfToImages(entry.file);
      for(const page of pages){
        const result=await Tesseract.recognize(page.blob,'eng',{});
        const text=fixOcrCommonErrors(result.data.text||'');
        combined += `\n===== ${entry.name}#${page.pageNo} =====\n` + text + '\n';
        const parsed=parseASNPageText(text, fallbackAsn);
        currentPageParses.push({source:`${entry.name}#${page.pageNo}`, asnNo:parsed.asnNo||fallbackAsn, rawLines:(parsed.rawLines||[]).map(r=>({...r, sourceName: entry.name}))});
      }
    }
  }
  $('ocrText').value=combined.trim();

  // auto build parsed result from page parses
  const rawLines=currentPageParses.flatMap(p=>p.rawLines);
  const uniqueAsns=[...new Set(currentPageParses.map(p=>p.asnNo).filter(Boolean))];
  const warnings=[];
  if(uniqueAsns.length !== currentFiles.length){
    warnings.push(`Đã quét ${currentFiles.length} file nhưng chỉ nhận ${uniqueAsns.length} ASN. Kiểm tra lại trang thiếu.`);
  }
  currentParsed={asnNo:uniqueAsns[0]||'', rawLines, warnings, meta:{pages:currentFiles.length||0}};
  currentParsed.results=processParsed(currentParsed);
  renderResults();
}

function parseCurrentText(){
  // Prefer per-file/page OCR results if available
  if(currentPageParses.length){
    const rawLines=currentPageParses.flatMap(p=>p.rawLines);
    const uniqueAsns=[...new Set(currentPageParses.map(p=>p.asnNo).filter(Boolean))];
    const warnings=[];
    if(uniqueAsns.length !== currentFiles.length){
      warnings.push(`Đã quét ${currentFiles.length} file nhưng chỉ nhận ${uniqueAsns.length} ASN. Kiểm tra lại trang thiếu.`);
    }
    currentParsed={asnNo:uniqueAsns[0]||'', rawLines, warnings, meta:{pages:currentFiles.length||0}};
  }else{
    currentParsed=parseASNPageText($('ocrText').value,'');
    currentParsed.meta={pages:currentFiles.length||1};
  }
  currentParsed.results=processParsed(currentParsed);
  renderResults();
}

function lookupSinglePacking(item,rev){
  const ni=normalizeItem(item), nr=normalizeRev(rev);
  const exact = state.singlePacking.find(p=>normalizeItem(p.item)===ni && normalizeRev(p.rev)===nr);
  if(exact) return exact;
  const matches = state.singlePacking.filter(p=>normalizeItem(p.item)===ni);
  return matches.length===1 ? matches[0] : null;
}
function findPairRule(item,rev){
  const ni=normalizeItem(item), nr=normalizeRev(rev);
  return state.pairPacking.find(p=>
    (normalizeItem(p.itemA)===ni && normalizeRev(p.revA)===nr) ||
    (normalizeItem(p.itemB)===ni && normalizeRev(p.revB)===nr)
  );
}

function processParsed(parsed){
  const grouped=new Map();
  (parsed.rawLines||[]).forEach(r=>{
    const key=`${r.asnNo}|${normalizeItem(r.itemNo)}|${normalizeRev(r.rev)}`;
    const prev=grouped.get(key);
    grouped.set(key,{
      type:'Single',
      asnNo:r.asnNo,
      item:normalizeItem(r.itemNo),
      rev:normalizeRev(r.rev),
      qty:(prev?.qty||0)+Number(r.quantity||0),
      packing:prev?.packing||0,
      carton:0, pcs:0, totalCarton:0,
      lineNo:[...new Set([...(prev?.lineNo?prev.lineNo.split(', ').filter(Boolean):[]), normalizeLineNo(r.lineNo||'')].filter(Boolean))].join(', '),
      remark:'',
      groupCode:r.groupCode||prev?.groupCode||'UNKNOWN',
      scheduleDate:r.scheduleDate||prev?.scheduleDate||'',
      scheduleTimes:[...new Set([...(prev?.scheduleTimes||[]), ...(r.scheduleTimes||[])])],
      sourceNames:[...new Set([...(prev?.sourceNames||[]), ...(r.sourceName?[r.sourceName]:[])])]
    });
  });
  const rows=Array.from(grouped.values());
  const byAsn={};
  rows.forEach(r=>{ (byAsn[r.asnNo]??=[]).push(r); });

  Object.values(byAsn).forEach(asnRows=>{
    asnRows.forEach(r=>{
      const pair=findPairRule(r.item,r.rev);
      if(pair){
        const oi=normalizeItem(pair.itemA)===normalizeItem(r.item)?normalizeItem(pair.itemB):normalizeItem(pair.itemA);
        const orv=normalizeItem(pair.itemA)===normalizeItem(r.item)?normalizeRev(pair.revB):normalizeRev(pair.revA);
        const mate=asnRows.find(x=>normalizeItem(x.item)===oi && normalizeRev(x.rev)===orv);
        if(mate){
          r.type='Pair';
          r.packing=Number(pair.packing||0);
          r.carton=r.packing?Math.floor(r.qty/r.packing):0;
          r.pcs=r.packing?r.qty%r.packing:0;
          r.totalCarton=r.packing?Math.ceil(r.qty/r.packing):0;
          r.remark='Pair';
          return;
        }
      }
      const single=lookupSinglePacking(r.item,r.rev);
      if((!r.packing||r.packing===0) && single) r.packing=Number(single.packing||0);
      r.type='Single';
      r.carton=r.packing?Math.floor(r.qty/r.packing):0;
      r.pcs=r.packing?r.qty%r.packing:0;
      r.totalCarton=r.packing?Math.ceil(r.qty/r.packing):0;
      r.remark=r.packing?'':'Thiếu packing';
    });
  });
  return rows;
}

function renderResults(){
  const tb=document.querySelector('#resultTable tbody');
  tb.innerHTML='';
  $('warningBox').classList.toggle('hidden', !(currentParsed.warnings||[]).length);
  $('warningBox').innerHTML=(currentParsed.warnings||[]).map(w=>`• ${w}`).join('<br>');
  currentParsed.results.forEach((r,i)=>{
    const tr=document.createElement('tr');
    tr.innerHTML=`<td>${i+1}</td>
<td><select class="cell-select" data-index="${i}" data-field="type"><option value="Single" ${r.type==='Single'?'selected':''}>Single</option><option value="Pair" ${r.type==='Pair'?'selected':''}>Pair</option></select></td>
<td><input class="cell-input mono" data-index="${i}" data-field="asnNo" value="${escapeHtml(r.asnNo)}"></td>
<td><input class="cell-input mono" data-index="${i}" data-field="item" value="${escapeHtml(r.item)}"></td>
<td><input class="cell-input mono" data-index="${i}" data-field="rev" value="${escapeHtml(r.rev)}"></td>
<td><input class="cell-input mono" data-index="${i}" data-field="qty" type="number" value="${r.qty}"></td>
<td><input class="cell-input mono" data-index="${i}" data-field="packing" type="number" value="${r.packing||0}"></td>
<td><span class="mono">${r.carton}</span></td>
<td><span class="mono">${r.pcs}</span></td>
<td><input class="cell-input mono" data-index="${i}" data-field="lineNo" value="${escapeHtml(r.lineNo||'')}"></td>
<td><input class="cell-input" data-index="${i}" data-field="remark" value="${escapeHtml(r.remark||'')}"></td>
<td><button class="small danger" data-del="${i}">Xóa</button></td>`;
    tb.appendChild(tr);
  });
  tb.querySelectorAll('[data-field]').forEach(inp=>{
    inp.addEventListener('change', syncRowsFromDOM);
    inp.addEventListener('blur', syncRowsFromDOM);
  });
  tb.querySelectorAll('[data-del]').forEach(btn=>btn.onclick=()=>{ currentParsed.results.splice(Number(btn.dataset.del),1); renderResults(); });
  updateResultMeta();
}

function syncRowsFromDOM(){
  document.querySelectorAll('#resultTable tbody [data-field]').forEach(inp=>{
    const idx=Number(inp.dataset.index), f=inp.dataset.field;
    let v=inp.type==='number'?Number(inp.value||0):inp.value;
    if(f==='item') v=normalizeItem(v);
    if(f==='rev') v=normalizeRev(v);
    if(f==='lineNo') v=normalizeLineNo(v);
    currentParsed.results[idx][f]=v;
  });
  recalcRows();
  updateResultMeta();
}
function recalcRows(){
  currentParsed.results.forEach(r=>{
    r.item=normalizeItem(r.item);
    r.rev=normalizeRev(r.rev);
    r.lineNo=normalizeLineNo(r.lineNo||'');
    const pair=findPairRule(r.item,r.rev);
    if(pair){
      const hasMate=currentParsed.results.find(x=>x!==r && (
        (normalizeItem(x.item)===normalizeItem(pair.itemA)&&normalizeRev(x.rev)===normalizeRev(pair.revA)) ||
        (normalizeItem(x.item)===normalizeItem(pair.itemB)&&normalizeRev(x.rev)===normalizeRev(pair.revB))
      ));
      if(hasMate){
        r.type='Pair'; r.packing=Number(pair.packing||0); r.carton=r.packing?Math.floor(r.qty/r.packing):0; r.pcs=r.packing?r.qty%r.packing:0; r.totalCarton=r.packing?Math.ceil(r.qty/r.packing):0; r.remark='Pair'; return;
      }
    }
    const single=lookupSinglePacking(r.item,r.rev);
    if((!r.packing||r.packing===0) && single) r.packing=Number(single.packing||0);
    r.type='Single'; r.carton=r.packing?Math.floor(r.qty/r.packing):0; r.pcs=r.packing?r.qty%r.packing:0; r.totalCarton=r.packing?Math.ceil(r.qty/r.packing):0; r.remark=r.packing?'':'Thiếu packing';
  });
}
function updateResultMeta(){
  const totalFull=currentParsed.results.reduce((a,r)=>a+Number(r.carton||0),0);
  const totalLooseQty=currentParsed.results.reduce((a,r)=>a+Number(r.pcs||0),0);
  const orderLoose=currentParsed.results.some(r=>Number(r.pcs||0)>0) ? 1 : 0;
  const totalCartons=totalFull + orderLoose;
  const miss=currentParsed.results.filter(r=>!Number(r.packing||0)).length;
  const asns=[...new Set(currentParsed.results.map(r=>r.asnNo).filter(Boolean))];
  $('resultMeta').textContent=`ASN: ${asns.join(', ')} | Số ASN: ${asns.length} | Pages: ${currentParsed.meta?.pages||0} | Dòng kết quả: ${currentParsed.results.length} | Total Cartons: ${totalCartons} | Loose PCS Qty: ${totalLooseQty} | Loose Cartons: ${orderLoose}${miss?' | Chưa có packing: '+miss:''}`;
}

function upsertSingle(item){
  const idx=state.singlePacking.findIndex(p=>normalizeItem(p.item)===normalizeItem(item.item)&&normalizeRev(p.rev)===normalizeRev(item.rev));
  if(idx>=0) state.singlePacking[idx]=item; else state.singlePacking.push(item);
  saveState(); renderPacking();
}
function upsertPair(item){
  const idx=state.pairPacking.findIndex(p=>normalizeItem(p.itemA)===normalizeItem(item.itemA)&&normalizeRev(p.revA)===normalizeRev(item.revA)&&normalizeItem(p.itemB)===normalizeItem(item.itemB)&&normalizeRev(p.revB)===normalizeRev(item.revB));
  if(idx>=0) state.pairPacking[idx]=item; else state.pairPacking.push(item);
  saveState(); renderPacking();
}
function renderPacking(){
  normalizePackingState(state);
  const singleQ = ($('singlePackingSearch')?.value || '').trim().toUpperCase();
  const pairQ = ($('pairPackingSearch')?.value || '').trim().toUpperCase();

  const sb=document.querySelector('#singleTable tbody'); sb.innerHTML='';
  state.singlePacking
    .slice()
    .sort((a,b)=>normalizeItem(a.item).localeCompare(normalizeItem(b.item)) || normalizeRev(a.rev).localeCompare(normalizeRev(b.rev)))
    .filter(p=>!singleQ || `${p.item} ${p.rev} ${p.note||''}`.toUpperCase().includes(singleQ))
    .forEach((p)=>{
      const key=`${normalizeItem(p.item)}|${normalizeRev(p.rev)}`;
      const tr=document.createElement('tr');
      tr.innerHTML=`<td><input class="cell-input mono pack-edit" data-kind="single" data-key="${key}" data-field="item" value="${p.item}"></td>
<td><input class="cell-input mono pack-edit" data-kind="single" data-key="${key}" data-field="rev" value="${p.rev}"></td>
<td><input class="cell-input mono pack-edit" data-kind="single" data-key="${key}" data-field="packing" type="number" value="${p.packing}"></td>
<td><input class="cell-input pack-edit" data-kind="single" data-key="${key}" data-field="note" value="${escapeHtml(p.note||'')}"></td>
<td><button class="small" data-save-single="${key}">Lưu</button> <button class="small danger" data-del-single="${key}">Xóa</button></td>`;
      sb.appendChild(tr);
    });
  sb.querySelectorAll('[data-save-single]').forEach(btn=>btn.onclick=()=>{
    const key=btn.dataset.saveSingle;
    const base=Array.from(sb.querySelectorAll(`.pack-edit[data-kind="single"][data-key="${key}"]`));
    const obj={};
    base.forEach(inp=>obj[inp.dataset.field]=inp.value);
    const oldIdx=state.singlePacking.findIndex(p=>`${normalizeItem(p.item)}|${normalizeRev(p.rev)}`===key);
    if(oldIdx>=0) state.singlePacking.splice(oldIdx,1);
    upsertSingle({item:normalizeItem(obj.item),rev:normalizeRev(obj.rev),packing:Number(obj.packing||0),note:obj.note||''});
  });
  sb.querySelectorAll('[data-del-single]').forEach(btn=>btn.onclick=()=>{
    const key=btn.dataset.delSingle;
    const idx=state.singlePacking.findIndex(p=>`${normalizeItem(p.item)}|${normalizeRev(p.rev)}`===key);
    if(idx>=0){ state.singlePacking.splice(idx,1); saveState(); renderPacking(); }
  });

  const pb=document.querySelector('#pairTable tbody'); pb.innerHTML='';
  state.pairPacking
    .slice()
    .sort((a,b)=>`${a.itemA}|${a.itemB}`.localeCompare(`${b.itemA}|${b.itemB}`))
    .filter(p=>!pairQ || `${p.itemA} ${p.revA} ${p.itemB} ${p.revB} ${p.note||''}`.toUpperCase().includes(pairQ))
    .forEach((p)=>{
      const key=`${p.itemA}|${p.revA}|${p.itemB}|${p.revB}`;
      const tr=document.createElement('tr');
      tr.innerHTML=`<td><input class="cell-input mono pair-edit" data-key="${key}" data-field="itemA" value="${p.itemA}"></td>
<td><input class="cell-input mono pair-edit" data-key="${key}" data-field="revA" value="${p.revA}"></td>
<td><input class="cell-input mono pair-edit" data-key="${key}" data-field="itemB" value="${p.itemB}"></td>
<td><input class="cell-input mono pair-edit" data-key="${key}" data-field="revB" value="${p.revB}"></td>
<td><input class="cell-input mono pair-edit" data-key="${key}" data-field="packing" type="number" value="${p.packing}"></td>
<td><input class="cell-input pair-edit" data-key="${key}" data-field="note" value="${escapeHtml(p.note||'')}"></td>
<td><button class="small" data-save-pair="${key}">Lưu</button> <button class="small danger" data-del-pair="${key}">Xóa</button></td>`;
      pb.appendChild(tr);
    });
  pb.querySelectorAll('[data-save-pair]').forEach(btn=>btn.onclick=()=>{
    const key=btn.dataset.savePair;
    const base=Array.from(pb.querySelectorAll(`.pair-edit[data-key="${key}"]`));
    const obj={};
    base.forEach(inp=>obj[inp.dataset.field]=inp.value);
    const oldIdx=state.pairPacking.findIndex(p=>`${p.itemA}|${p.revA}|${p.itemB}|${p.revB}`===key);
    if(oldIdx>=0) state.pairPacking.splice(oldIdx,1);
    upsertPair({itemA:normalizeItem(obj.itemA),revA:normalizeRev(obj.revA),itemB:normalizeItem(obj.itemB),revB:normalizeRev(obj.revB),packing:Number(obj.packing||0),note:obj.note||''});
  });
  pb.querySelectorAll('[data-del-pair]').forEach(btn=>btn.onclick=()=>{
    const key=btn.dataset.delPair;
    const idx=state.pairPacking.findIndex(p=>`${p.itemA}|${p.revA}|${p.itemB}|${p.revB}`===key);
    if(idx>=0){ state.pairPacking.splice(idx,1); saveState(); renderPacking(); }
  });
}
async function importPackingExcel(e){
  const file=e.target.files?.[0];
  if(!file) return;
  const data=await file.arrayBuffer();
  const wb=XLSX.read(data,{type:'array'});
  const ws=wb.Sheets[wb.SheetNames[0]];
  const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
  if(!rows.length) return alert('File Excel trống.');
  const header=rows[0].map(x=>String(x).trim().toLowerCase());
  let itemIdx=header.findIndex(h=>h.includes('mã')||h.includes('ma')||h.includes('item'));
  let revIdx=header.findIndex(h=>h.includes('rev'));
  let packingIdx=header.findIndex(h=>h.includes('packing'));
  if(itemIdx<0) itemIdx=0;
  if(revIdx<0) revIdx=1;
  if(packingIdx<0) packingIdx=2;
  let added=0, updated=0;
  rows.slice(1).forEach(r=>{
    const item=normalizeItem(r[itemIdx]); const rev=normalizeRev(r[revIdx]); const packing=Number(r[packingIdx]||0);
    if(!/^[35]\d{8}$/.test(item)||!rev||!packing) return;
    const idx=state.singlePacking.findIndex(p=>normalizeItem(p.item)===item&&normalizeRev(p.rev)===rev);
    if(idx>=0){ state.singlePacking[idx]={...state.singlePacking[idx], item, rev, packing, note:'Imported Excel'}; updated++; }
    else { state.singlePacking.push({item, rev, packing, note:'Imported Excel'}); added++; }
  });
  saveState(); renderPacking(); e.target.value='';
  alert(`Đã import Packing Excel.\nThêm mới: ${added}\nCập nhật: ${updated}`);
}
function exportPackingExcel(){
  const wb=XLSX.utils.book_new();
  const ws1=XLSX.utils.aoa_to_sheet([['Mã Hàng','Rev','Packing','Note'], ...state.singlePacking.map(p=>[p.item,p.rev,p.packing,p.note||''])]);
  const ws2=XLSX.utils.aoa_to_sheet([['Item A','Rev A','Item B','Rev B','Packing','Note'], ...state.pairPacking.map(p=>[p.itemA,p.revA,p.itemB,p.revB,p.packing,p.note||''])]);
  XLSX.utils.book_append_sheet(wb,ws1,'SinglePacking');
  XLSX.utils.book_append_sheet(wb,ws2,'PairPacking');
  XLSX.writeFile(wb,'PACKING_MASTER_EXPORT.xlsx');
}

function groupResultsByASN(rows){
  const g={};
  rows.forEach(r=>{
    if(!g[r.asnNo]) g[r.asnNo]={rows:[],scheduleDate:r.scheduleDate||'',scheduleTimes:r.scheduleTimes||[],groupCode:r.groupCode||'UNKNOWN'};
    g[r.asnNo].rows.push({...r});
    g[r.asnNo].sourceNames=[...new Set([...(g[r.asnNo].sourceNames||[]), ...(r.sourceNames||[])])];
    g[r.asnNo].scheduleDate=g[r.asnNo].scheduleDate||r.scheduleDate||'';
    g[r.asnNo].scheduleTimes=[...new Set([...(g[r.asnNo].scheduleTimes||[]), ...(r.scheduleTimes||[])])];
    g[r.asnNo].groupCode=g[r.asnNo].groupCode||r.groupCode||'UNKNOWN';
  });
  return g;
}
async function fileToDataURL(file){
  return new Promise((resolve,reject)=>{
    const reader=new FileReader();
    reader.onload=()=>resolve(reader.result);
    reader.onerror=reject;
    reader.readAsDataURL(file);
  });
}

async function saveCurrentASN(){
  if(!currentParsed.results.length) return alert('Chưa có dữ liệu để lưu.');
  const groups=groupResultsByASN(currentParsed.results);
  const now=new Date();
  const batchId=crypto.randomUUID();
  const label=now.toLocaleTimeString([], {hour:'2-digit', minute:'2-digit'})+' '+now.toLocaleDateString();

  const batchEntries=[];
  for(const [asnNo,group] of Object.entries(groups)){
    const totalFull = group.rows.reduce((a,r)=>a+Number(r.carton||0),0);
    const looseCarton = group.rows.some(r=>Number(r.pcs||0)>0) ? 1 : 0;
    const looseQty = group.rows.reduce((a,r)=>a+Number(r.pcs||0),0);
    const totals={cartons: totalFull + looseCarton, pcs: looseCarton, pcsQty: looseQty};

    let matchedFiles=[];
    const sourceBaseNames=[...new Set((group.sourceNames||[]).map(s=>String(s).split('#')[0]))];
    if(sourceBaseNames.length){
      matchedFiles=currentFiles.filter(entry=>sourceBaseNames.includes(entry.name));
    }
    if(!matchedFiles.length){
      const byAsn=currentFiles.filter(entry=>entry.name.toUpperCase().includes(asnNo.toUpperCase()));
      if(byAsn.length) matchedFiles=byAsn;
    }
    if(!matchedFiles.length){
      matchedFiles=[...currentFiles];
    }

    const originalFiles=[];
    for(const entry of matchedFiles){
      try{
        originalFiles.push({name:entry.name,type:entry.type,dataUrl:await fileToDataURL(entry.file)});
      }catch(err){
        console.error('fileToDataURL failed', err);
      }
    }

    batchEntries.push({
      id:crypto.randomUUID(),
      asnNo,
      pages:matchedFiles.length||currentFiles.length,
      totals,
      results:group.rows,
      originalFiles,
      hasOriginalFiles: originalFiles.length>0,
      scheduleDate:group.scheduleDate,
      scheduleTimes:group.scheduleTimes,
      groupCode:group.groupCode||'UNKNOWN',
      sourceNames:group.sourceNames||[]
    });
  }

  const batchTotals=batchEntries.reduce((a,e)=>({
    cartons:a.cartons+Number(e.totals?.cartons||0),
    pcs:a.pcs+Number(e.totals?.pcs||0),
    pcsQty:a.pcsQty+Number(e.totals?.pcsQty||0)
  }),{cartons:0,pcs:0,pcsQty:0});

  state.historyBatches = state.historyBatches || [];
  state.historyBatches.unshift({
    id:batchId,
    label,
    createdAt:now.toISOString(),
    pages:currentFiles.length||0,
    totals:batchTotals,
    asnList:batchEntries.map(e=>e.asnNo),
    entries:batchEntries
  });

  try{
    saveState();
    renderHistory();
    alert('Đã lưu batch ASN vào History.');
  }catch(err){
    console.error(err);
    alert('Lưu History thất bại.');
  }
}

function renderHistory(){
  const q=($('historySearch').value||'').trim().toLowerCase();
  const tb=document.querySelector('#historyTable tbody');
  tb.innerHTML='';
  const batches=(state.historyBatches||[]);
  batches
    .filter(b=>{
      if(!q) return true;
      return (b.label||'').toLowerCase().includes(q) || (b.asnList||[]).some(a=>a.toLowerCase().includes(q));
    })
    .forEach(b=>{
      const d=new Date(b.createdAt);
      const tr=document.createElement('tr');
      tr.innerHTML=`<td>${d.toLocaleDateString()}</td><td>${d.toLocaleTimeString([], {hour:'2-digit', minute:'2-digit'})}</td><td class="mono">${b.label||''}</td><td>${b.pages||0}</td><td>${b.totals?.cartons||0}</td><td>${b.totals?.pcs||0}</td><td><button class="small history-open-btn" type="button">Mở</button></td>`;
      tr.querySelector('.history-open-btn').addEventListener('click', ()=>renderHistoryBatchDetail(b.id));
      tb.appendChild(tr);
    });
}
function renderHistoryBatchDetail(batchId){
  const b=(state.historyBatches||[]).find(x=>x.id===batchId);
  if(!b) return;
  $('historyDetail').classList.remove('hidden');
  $('historyDetail').innerHTML=`<h2>🗂 Batch ${b.label||''}</h2>
  <div style="margin:6px 0 12px 0;color:#475569;">Lưu lúc: ${new Date(b.createdAt).toLocaleString()} | ASN: ${b.asnList.length} | Pages: ${b.pages||0}</div>
  <div class="toolbar wrap">
    <button id="exportBatchSavedBtn">📊 Xuất Excel Batch</button>
  </div>
  <div class="table-wrap">
    <table>
      <thead><tr><th>ASN</th><th>Tổng thùng</th><th>Thùng lẻ</th><th>Pages</th><th></th></tr></thead>
      <tbody>
        ${(b.entries||[]).map(e=>`<tr>
          <td class="mono">${e.asnNo}</td>
          <td>${e.totals?.cartons||0}</td>
          <td>${e.totals?.pcs||0}</td>
          <td>${e.pages||0}</td>
          <td><button class="small batch-asn-open" data-asn="${e.asnNo}">Mở ASN</button></td>
        </tr>`).join('')}
      </tbody>
    </table>
  </div>
  <div id="batchAsnDetail"></div>`;
  $('exportBatchSavedBtn').onclick=()=>{
    const groups={};
    (b.entries||[]).forEach(e=>groups[e.asnNo]={rows:e.results,scheduleDate:e.scheduleDate,scheduleTimes:e.scheduleTimes,groupCode:e.groupCode});
    exportMultiASNWorkbook(groups);
  };
  document.querySelectorAll('.batch-asn-open').forEach(btn=>{
    btn.onclick=()=>{
      const entry=(b.entries||[]).find(x=>x.asnNo===btn.dataset.asn);
      if(!entry) return;
      const originalFiles = entry.originalFiles || [];
      const box=$('batchAsnDetail');
      box.innerHTML=`<div class="card" style="margin-top:12px;"><h2>📄 ASN ${entry.asnNo}</h2>
      <div class="toolbar wrap">
        <button id="previewOriginalBtn">👁 Xem đơn</button>
        <button id="downloadOriginalBtn" class="ghost">⬇ Lưu file gốc</button>
        <button id="exportDetailBtn" class="ghost">📊 Lưu Excel .xlsx</button>
      </div>
      <div class="table-wrap">
        <table>
          <thead><tr><th>STT</th><th>Type</th><th>Item</th><th>Rev</th><th>Qty</th><th>Packing</th><th>Carton</th><th>PCS</th><th>Line</th><th>Remark</th></tr></thead>
          <tbody>${entry.results.map((r,i)=>`<tr><td>${i+1}</td><td>${r.type}</td><td class="mono">${r.item}</td><td class="mono">${r.rev}</td><td class="mono">${r.qty}</td><td class="mono">${r.packing||''}</td><td class="mono">${r.carton||0}</td><td class="mono">${r.pcs||0}</td><td class="mono">${escapeHtml(r.lineNo||'')}</td><td>${escapeHtml(r.remark||'')}</td></tr>`).join('')}</tbody>
        </table>
      </div></div>`;
      $('previewOriginalBtn').onclick=()=>previewStoredFiles(originalFiles,entry.asnNo);
      $('downloadOriginalBtn').onclick=()=>downloadStoredFiles(originalFiles,entry.asnNo);
      $('exportDetailBtn').onclick=()=>exportMultiASNWorkbook({[entry.asnNo]:{rows:entry.results,scheduleDate:entry.scheduleDate,scheduleTimes:entry.scheduleTimes,groupCode:entry.groupCode}});
    };
  });
}
function renderHistoryDetail(id){
  // backward compatibility: if old single entry exists
  const h=(state.history||[]).find(x=>x.id===id);
  if(!h) return;
  $('historyDetail').classList.remove('hidden');
  $('historyDetail').innerHTML=`<h2>📄 ASN ${h.asnNo}</h2>
  <div class="toolbar wrap">
    <button id="previewOriginalBtn">👁 Xem đơn</button>
    <button id="downloadOriginalBtn" class="ghost">⬇ Lưu file gốc</button>
    <button id="exportDetailBtn" class="ghost">📊 Lưu Excel .xlsx</button>
  </div>
  <div class="table-wrap">
    <table>
      <thead><tr><th>STT</th><th>Type</th><th>Item</th><th>Rev</th><th>Qty</th><th>Packing</th><th>Carton</th><th>PCS</th><th>Line</th><th>Remark</th></tr></thead>
      <tbody>${h.results.map((r,i)=>`<tr><td>${i+1}</td><td>${r.type}</td><td class="mono">${r.item}</td><td class="mono">${r.rev}</td><td class="mono">${r.qty}</td><td class="mono">${r.packing||''}</td><td class="mono">${r.totalCarton||0}</td><td class="mono">${r.pcs||0}</td><td class="mono">${escapeHtml(r.lineNo||'')}</td><td>${escapeHtml(r.remark||'')}</td></tr>`).join('')}</tbody>
    </table>
  </div>`;
  $('previewOriginalBtn').onclick=()=>previewStoredFiles(h.originalFiles||[],h.asnNo);
  $('downloadOriginalBtn').onclick=()=>downloadStoredFiles(h.originalFiles||[],h.asnNo);
  $('exportDetailBtn').onclick=async ()=>{ await exportMultiASNWorkbook({[h.asnNo]:{rows:h.results,scheduleDate:h.scheduleDate,scheduleTimes:h.scheduleTimes,groupCode:h.groupCode}}); };
}
function dataUrlToFile(dataUrl, filename, mimeHint='application/octet-stream'){
  const parts=String(dataUrl||'').split(',');
  if(parts.length<2) return null;
  const mimeMatch=parts[0].match(/data:(.*?);base64/);
  const mime=mimeMatch?mimeMatch[1]:mimeHint;
  const binary=atob(parts[1]); const len=binary.length; const bytes=new Uint8Array(len);
  for(let i=0;i<len;i++) bytes[i]=binary.charCodeAt(i);
  return new File([bytes], filename, {type:mime});
}
async function shareFilesIfPossible(files, title='OCR ASN GOODMARK FINAL'){
  try{
    if(navigator.share && files?.length && navigator.canShare && navigator.canShare({files})){
      await navigator.share({files, title});
      return true;
    }
  }catch(err){ console.error(err); }
  return false;
}
function downloadBlobFallback(blob, filename){
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url; a.download=filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(()=>URL.revokeObjectURL(url),1500);
}
async function saveBlobOrShare(blob, filename, mime='application/octet-stream', title='OCR ASN GOODMARK FINAL'){
  const file=new File([blob], filename, {type:mime});
  const shared=await shareFilesIfPossible([file], title);
  if(shared) return true;
  downloadBlobFallback(blob, filename);
  return true;
}
async function downloadStoredFiles(files,asnNo){
  if(!files?.length) return alert('Không có file gốc.');
  const prepared=[];
  files.forEach((f,idx)=>{
    const ext=(f.name||'file').includes('.') ? f.name.split('.').pop() : (((f.type||'').includes('pdf')) ? 'pdf' : 'jpg');
    const filename=files.length===1 ? `ASN_${asnNo}.${ext}` : `ASN_${asnNo}_${idx+1}.${ext}`;
    const fileObj=dataUrlToFile(f.dataUrl, filename, f.type||'application/octet-stream');
    if(fileObj) prepared.push(fileObj);
  });
  if(prepared.length){
    const shared=await shareFilesIfPossible(prepared, `ASN ${asnNo}`);
    if(shared) return;
    prepared.forEach(file=>downloadBlobFallback(file, file.name));
    return;
  }
  alert('Không xử lý được file gốc.');
}
function previewStoredFiles(files,asnNo){
  if(!files?.length) return alert('Không có file gốc để xem.');
  const modal=$('previewModal'), content=$('previewContent');
  if(!modal || !content) return alert('Preview chưa sẵn sàng.');
  content.innerHTML=`<h2 style="color:#fff;margin:0 0 8px 0;">Đơn hàng ${asnNo}</h2>`;
  files.forEach((f,idx)=>{
    const wrap=document.createElement('div');
    wrap.style.cssText='background:#fff;border-radius:14px;padding:12px;';
    const title=document.createElement('div');
    title.style.cssText='font-weight:700;margin-bottom:8px;word-break:break-word;';
    title.textContent=f.name||`File ${idx+1}`;
    wrap.appendChild(title);
    if(String(f.type||'').includes('pdf') || String(f.name||'').toLowerCase().endsWith('.pdf')){
      const iframe=document.createElement('iframe');
      iframe.src=f.dataUrl;
      iframe.style.cssText='width:100%;height:70vh;border:1px solid #d1d5db;border-radius:8px;background:#fff;';
      wrap.appendChild(iframe);
    }else{
      const img=document.createElement('img');
      img.src=f.dataUrl;
      img.style.cssText='width:100%;height:auto;border-radius:8px;display:block;background:#f3f4f6;';
      wrap.appendChild(img);
    }
    content.appendChild(wrap);
  });
  modal.classList.remove('hidden');
}

async function exportMultiASNWorkbook(groups){
  function borderAll(color='000000'){ return {top:{style:'thin',color:{rgb:color}},bottom:{style:'thin',color:{rgb:color}},left:{style:'thin',color:{rgb:color}},right:{style:'thin',color:{rgb:color}}};}
  function cell(v,s){ return {v,t:typeof v==='number'?'n':'s',s};}
  function headerBarStyle(){ return {font:{bold:true,sz:16,color:{rgb:'FFFFFF'}},fill:{fgColor:{rgb:'1F4E78'}},alignment:{horizontal:'center',vertical:'center'},border:borderAll('1F1F1F')};}
  function infoStyle(){ return {font:{bold:false,sz:11},alignment:{horizontal:'left',vertical:'center'},border:borderAll('A0A0A0')};}
  function tableHeaderStyle(){ return {font:{bold:true,sz:11,color:{rgb:'000000'}},fill:{fgColor:{rgb:'D9E2F3'}},alignment:{horizontal:'center',vertical:'center'},border:borderAll('000000')};}
  function bodyStyle(idx,center=true){ return {font:{sz:11},fill:{fgColor:{rgb:idx%2===0?'FFFFFF':'F8FAFC'}},alignment:{horizontal:center?'center':'left',vertical:'center'},border:borderAll('7F7F7F')};}
  function totalStyle(){ return {font:{bold:true,sz:11},fill:{fgColor:{rgb:'FFF2CC'}},alignment:{horizontal:'center',vertical:'center'},border:borderAll('000000')};}

  function buildSheet(asnEntries){
    const ws={}; ws['!merges']=[]; let row=1;
    asnEntries.forEach(([asnNo,g])=>{
      const totalFull = g.rows.reduce((a,r)=>a+Number(r.carton||0),0);
      const looseCarton = g.rows.some(r=>Number(r.pcs||0)>0) ? 1 : 0;
      const orderTotalCartons = totalFull + looseCarton;

      ws[`A${row}`]=cell(`${asnNo}`, headerBarStyle());
      ws[`A${row+1}`]=cell(`Date: ${g.scheduleDate||'N/A'}`, infoStyle());
      ws[`A${row+2}`]=cell(`Time: ${(g.scheduleTimes||[]).join(', ')||'N/A'}`, infoStyle());
      const headers=['STT','Type','ASN','Item','Rev','Qty','Packing','Carton','PCS','Line'];
      headers.forEach((h,i)=>{ ws[XLSX.utils.encode_cell({r:row+3-1,c:i})]=cell(h, tableHeaderStyle()); });
      g.rows.forEach((r,idx)=>{
        const vals=[idx+1,r.type,asnNo,r.item,r.rev,Number(r.qty||0),Number(r.packing||0),Number(r.carton||0),Number(r.pcs||0),r.lineNo||''];
        vals.forEach((v,c)=>{ ws[XLSX.utils.encode_cell({r:row+4+idx-1,c})]=cell(v, bodyStyle(idx, c!==9)); });
      });
      const totalRow=row+4+g.rows.length;
      ws[`A${totalRow}`]=cell('Tổng ASN', totalStyle());
      ws[`H${totalRow}`]=cell(orderTotalCartons, totalStyle());
      ws['!merges'].push(
        XLSX.utils.decode_range(`A${row}:J${row}`),
        XLSX.utils.decode_range(`A${totalRow}:G${totalRow}`),
        XLSX.utils.decode_range(`H${totalRow}:I${totalRow}`)
      );
      row=totalRow+2;
    });
    ws['!cols']=[{wch:6},{wch:8},{wch:14},{wch:14},{wch:8},{wch:10},{wch:12},{wch:10},{wch:8},{wch:18}];
    ws['!ref']=`A1:J${row}`;
    return ws;
  }

  const wb=XLSX.utils.book_new();
  const allEntries=Object.entries(groups);
  const xc2Entries=allEntries.filter(([,g])=>(g.groupCode||'').toUpperCase()==='XC2-TC2');
  const xc5Entries=allEntries.filter(([,g])=>(g.groupCode||'').toUpperCase()==='XC5-TC5');
  const otherEntries=allEntries.filter(([,g])=>!['XC2-TC2','XC5-TC5'].includes((g.groupCode||'').toUpperCase()));

  if(xc2Entries.length) XLSX.utils.book_append_sheet(wb, buildSheet(xc2Entries), 'XC2-TC2');
  if(xc5Entries.length) XLSX.utils.book_append_sheet(wb, buildSheet(xc5Entries), 'XC5-TC5');
  if(otherEntries.length) XLSX.utils.book_append_sheet(wb, buildSheet(otherEntries), 'OTHER');
  if(!wb.SheetNames.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['No data']]), 'ASN Report');

  XLSX.writeFile(wb, `ASN_BATCH_${new Date().toISOString().replace(/[-:T]/g,'').slice(0,15)}.xlsx`);
}
async function exportBatchWorkbook(){
  if(!state.history.length) return alert('History chưa có ASN nào.');
  const groups={};
  state.history.forEach(h=>groups[h.asnNo]={rows:h.results,scheduleDate:h.scheduleDate,scheduleTimes:h.scheduleTimes});
  await exportMultiASNWorkbook(groups);
}

function escapeHtml(str){ return String(str).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;'); }

function init(){
  document.querySelectorAll('.nav-btn').forEach(btn=>btn.onclick=()=>{
    document.querySelectorAll('.nav-btn').forEach(b=>b.classList.remove('active'));
    document.querySelectorAll('.tab-page').forEach(p=>p.classList.remove('active'));
    btn.classList.add('active'); $(btn.dataset.tab).classList.add('active');
  });
  $('imageInput').onchange=handleFiles; $('fileInput').onchange=handleFiles;
  $('clearScanBtn').onclick=clearScanState; $('sampleBtn').onclick=loadSampleText; $('parseBtn').onclick=parseCurrentText;
  $('addRowBtn').onclick=()=>{ currentParsed.results.push({type:'Single',asnNo:'',item:'',rev:'01',qty:0,packing:0,carton:0,pcs:0,totalCarton:0,lineNo:'',remark:''}); renderResults(); };
  $('recalcBtn').onclick=()=>{ syncRowsFromDOM(); recalcRows(); renderResults(); };
  $('saveAsnBtn').onclick=saveCurrentASN;
  $('exportExcelBtn').onclick=async ()=>{ await exportMultiASNWorkbook(groupResultsByASN(currentParsed.results)); };
  $('exportBatchBtn').onclick=async ()=>{ await exportBatchWorkbook(); };
  $('historySearch').oninput=renderHistory;
  $('showSingleBtn').onclick=()=>setPackingTab('single'); $('showPairBtn').onclick=()=>setPackingTab('pair');
  if($('singlePackingSearch')) $('singlePackingSearch').oninput=renderPacking;
  if($('pairPackingSearch')) $('pairPackingSearch').oninput=renderPacking;
  $('singleForm').onsubmit=e=>{ e.preventDefault(); const f=new FormData(e.target); upsertSingle({item:normalizeItem(f.get('item')),rev:normalizeRev(f.get('rev')),packing:Number(f.get('packing')),note:f.get('note')||''}); e.target.reset(); };
  $('pairForm').onsubmit=e=>{ e.preventDefault(); const f=new FormData(e.target); upsertPair({itemA:normalizeItem(f.get('itemA')),revA:normalizeRev(f.get('revA')),itemB:normalizeItem(f.get('itemB')),revB:normalizeRev(f.get('revB')),packing:Number(f.get('packing')),note:f.get('note')||''}); e.target.reset(); };
  $('packingImportInput').onchange=importPackingExcel; $('exportPackingBtn').onclick=exportPackingExcel;
  $('runOcrBtn').onclick=async ()=>{ await runOcrOnFiles(); }; $('scanQrBtn').onclick=async ()=>{ await scanQrFromImages(); };
  if($('previewCloseBtn')) $('previewCloseBtn').onclick=()=>{ $('previewModal').classList.add('hidden'); $('previewContent').innerHTML=''; };
  if($('previewModal')) $('previewModal').addEventListener('click',e=>{ if(e.target.id==='previewModal'){ $('previewModal').classList.add('hidden'); $('previewContent').innerHTML=''; } });
  renderPacking(); renderHistory();
}

init();

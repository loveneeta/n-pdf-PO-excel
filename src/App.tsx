import { useMemo, useState } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from './download';
import * as pdfjsLib from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import Tesseract from 'tesseract.js';
import { Upload, FileText, FileSpreadsheet, Download, Trash2, AlertTriangle, CheckCircle2, LoaderCircle, Files } from 'lucide-react';

(pdfjsLib as any).GlobalWorkerOptions.workerSrc = workerUrl;

type Row = { SO:string; QTY:string; ITEM:string; L:string; W:string; H:string; INCH:string; PLY:string; TYPE:string; source:string; confidence:'high'|'review'; [key:string]: any };
const HEADERS = ['SO','QTY','ITEM','L','W','H','INCH','PLY','TYPE'] as const;
const CUSTOMER_NAME = 'Noble Art And Craft House';
const IMPORT_HEADERS = ['Customer Name','SO/NO (Sales Orders)','SO Quantities (Multi Qty e.g. 600, 400)','Item Code','Dimension Unit','Length','Width','Height','Ply / Board','Flute Type','Joint (inch)','Deckle Margin (inch)','Flap Type','Overlap Length (inch)','Total Quantity','Wastage %','Setup Cost','Tooling Cost','Printing Required (Yes/No)','Printing / Box','Punching / Box','Lamination / Box','Pasting / Box','Transport / Box','Other / Box','Factory Overhead %','Paper Combo Name'] as const;
const DEFAULTS: Record<string, string | number> = {
  'Dimension Unit':'inch','Flute Type':'B/C','Joint (inch)':2,'Deckle Margin (inch)':2,
  'Overlap Length (inch)':1,'Wastage %':6,'Setup Cost':0,'Tooling Cost':0,
  'Printing Required (Yes/No)':'yes','Printing / Box':0,'Punching / Box':0,
  'Lamination / Box':0,'Pasting / Box':0,'Transport / Box':0,'Other / Box':0,
  'Factory Overhead %':5,
  'Paper Combo Name':'18/140+18/100+18/140+18/100+18/140+18/100+18/140'
};
const blank = (source:string):Row => ({SO:'',QTY:'',ITEM:'',L:'',W:'',H:'',INCH:'INCH',PLY:'',TYPE:'',source,confidence:'review'});
const clean = (s:string) => s.replace(/\s+/g,' ').trim();
const getTemplateValues = (r:Row):Record<string,string|number> => {
  const values:Record<string,string|number> = {
    'Customer Name': CUSTOMER_NAME,
    'SO/NO (Sales Orders)': (r['SO/NO (Sales Orders)'] ?? r.SO.replace(/^SO\//i,'')),
    'SO Quantities (Multi Qty e.g. 600, 400)': (r['SO Quantities (Multi Qty e.g. 600, 400)'] ?? r.QTY),
    'Item Code': r['Item Code'] ?? r.ITEM,
    'Dimension Unit': r['Dimension Unit'] ?? r.INCH ?? 'inch',
    'Length': r.Length ?? r.L, 'Width': r.Width ?? r.W, 'Height': r.Height ?? r.H,
    'Ply / Board': r['Ply / Board'] ?? r.PLY,
    'Flap Type': r['Flap Type'] ?? (/^(DF|SF)$/i.test(r.TYPE) ? r.TYPE.toUpperCase() : ''),
    'Total Quantity': r['Total Quantity'] ?? r.QTY
  };
  for (const [key,value] of Object.entries(DEFAULTS)) {
    if (values[key] === undefined || values[key] === null || values[key] === '') values[key]=value;
  }
  for (const h of IMPORT_HEADERS) {
    if (r[h] !== undefined) values[h]=r[h];
  }
  return values;
};

const dimPattern = /(\d+(?:\.\d+)?)\s*(?:x|×|\*)\s*(\d+(?:\.\d+)?)\s*(?:x|×|\*)\s*(\d+(?:\.\d+)?)/i;

function parseText(text:string, fileName:string):Row[] {
  // PDF.js frequently splits a visual table row across many text spans.
  const normalized = text.replace(/\u00a0/g, ' ').replace(/[“”]/g, '"').replace(/[×✕]/g, 'x').replace(/\s+/g, ' ').trim();
  const so = (normalized.match(/\bSO\s*\/\s*\d{6,}\b/i)?.[0] || normalized.match(/\bSO\s*\/\s*[A-Z0-9-]+\b/i)?.[0] || '').replace(/\s+/g, '');
  const headerRe = /\[?\s*Box\s*-\s*\d+\s*\]?/ig;
  const matches = [...normalized.matchAll(headerRe)];
  const chunks:string[] = [];
  if (matches.length) {
    for (let i=0;i<matches.length;i++) {
      const from=matches[i].index ?? 0;
      const to=i+1<matches.length ? (matches[i+1].index ?? normalized.length) : normalized.length;
      chunks.push(normalized.slice(from,to));
    }
  } else chunks.push(...text.split(/(?=\b\d+(?:,\d{3})*(?:\.\d+)?\s*PCS\b)/i).filter(Boolean));

  const rows:Row[]=[];
  for (const chunk of chunks) {
    const qtyMatch=chunk.match(/\b(\d+(?:,\d{3})*(?:\.\d+)?)\s*(?:PCS|PC|PIECES)\b/i);
    const itemMatch=chunk.match(/\bITEM\s*:\s*([A-Z0-9-]+)/i);
    const plyMatch=chunk.match(/\b([357])\s*[- ]?\s*PLY\b/i);
    const sizeMatch=chunk.match(/\bSize\s*:\s*(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)/i);
    const allDims=[...chunk.matchAll(/(\d+(?:\.\d+)?)\s*(?:x|\*)\s*(\d+(?:\.\d+)?)\s*(?:x|\*)\s*(\d+(?:\.\d+)?)/ig)];
    let type='';
    // Prefer the explicit DF/SF suffix printed in the reference field.
    const suffixType = chunk.match(/\[\s*\d+(?:\.\d+)?\s*\]\s*["']?\s*(DF|SF)\b/i)
      || chunk.match(/["']\s*(DF|SF)\s*["']/i);
    if (suffixType) type = suffixType[1].toUpperCase();
    else if (/Triangle\s*Box\s*\(\s*B\s*\)/i.test(chunk)) type='Triangle Box (B)';
    else if (/Triangle\s*Box\s*\(\s*C\s*\)/i.test(chunk)) type='Triangle Box (C)';
    else if (/Triangle\s*Box/i.test(chunk)) type=(chunk.match(/Triangle\s*Box\s*\([^)]*\)/i)?.[0] || 'Triangle Box');
    else if (/\bSF\b/i.test(chunk)) type='SF';
    const dims=sizeMatch || allDims[0];
    const row=blank(fileName);
    row.SO=so; row.QTY=qtyMatch ? qtyMatch[1].replace(/,/g,'').replace(/\.000$/,'') : '';
    row.ITEM=itemMatch ? itemMatch[1].toUpperCase() : '';
    row.L=dims?.[1] || ''; row.W=dims?.[2] || ''; row.H=dims?.[3] || '';
    row.INCH='INCH'; row.PLY=plyMatch ? plyMatch[1] : ''; row.TYPE=type;
    row.confidence=(row.SO && row.QTY && row.ITEM && row.L && row.W && row.H && row.PLY && row.TYPE) ? 'high' : 'review';
    if (row.QTY || row.ITEM || row.PLY || row.L) rows.push(row);
  }
  const unique=new Map<string,Row>();
  rows.forEach(r=>unique.set([r.source,r.SO,r.ITEM,r.QTY,r.L,r.W,r.H,r.PLY,r.TYPE].join('|'),r));
  return [...unique.values()];
}

async function pdfText(file:File, setStatus:(s:string)=>void):Promise<string> {
  const doc = await (pdfjsLib as any).getDocument({data: await file.arrayBuffer()}).promise;
  let all = '';
  for (let p=1;p<=doc.numPages;p++) {
    setStatus(`${file.name}: reading page ${p} of ${doc.numPages}`);
    const page = await doc.getPage(p);
    const content = await page.getTextContent();
    const pageText = content.items.map((x:any)=>x.str || '').join('\n');
    if (pageText.replace(/\s/g,'').length > 30) all += '\n' + pageText;
    else {
      setStatus(`${file.name}: OCR page ${p} of ${doc.numPages}`);
      const viewport = page.getViewport({scale:1.8});
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width; canvas.height = viewport.height;
      const ctx = canvas.getContext('2d')!;
      await page.render({canvasContext:ctx, viewport}).promise;
      const result = await Tesseract.recognize(canvas, 'eng');
      all += '\n' + result.data.text;
    }
  }
  return all;
}

export default function App() {
  const [files,setFiles] = useState<File[]>([]);
  const [rows,setRows] = useState<Row[]>([]);
  const [status,setStatus] = useState('');
  const [busy,setBusy] = useState(false);
  const [done,setDone] = useState(false);
  const reviewCount = useMemo(()=>rows.filter(r=>r.confidence==='review').length,[rows]);
  const addFiles = (list:FileList|null) => { if(!list)return; setFiles(prev=>[...prev,...Array.from(list).filter(f=>f.type==='application/pdf'||f.name.toLowerCase().endsWith('.pdf'))]); setDone(false); };
  const convert = async () => {
    if(!files.length)return;
    setBusy(true); setRows([]); setDone(false);
    try {
      const all:Row[]=[];
      for(const file of files) {
        const text=await pdfText(file,setStatus);
        const parsed=parseText(text,file.name);
        if(!parsed.length) all.push({...blank(file.name),confidence:'review'});
        else all.push(...parsed);
      }
      setRows(all); setDone(true); setStatus(`Finished: ${all.length} item rows extracted.`);
    } catch(e:any) { setStatus(`Conversion failed: ${e?.message || 'Unknown error'}`); }
    finally { setBusy(false); }
  };
  const updateRow = (idx:number,key:string,value:string) => setRows(prev=>prev.map((r,i)=>i===idx?{...r,[key]:value,confidence:'review'}:r));
  const makeWorkbook = async (data:Row[], title:string) => {
    const wb=new ExcelJS.Workbook(); wb.creator='PDF to Excel Converter';
    const ws=wb.addWorksheet('Bulk Calculation Template');
    ws.addRow([...IMPORT_HEADERS]);
    data.forEach(r => { const values=getTemplateValues(r); ws.addRow(IMPORT_HEADERS.map(h => values[h] ?? '')); });
    ws.views=[{state:'frozen',ySplit:1}];
    ws.autoFilter={from:'A1',to:`AA${data.length+1}`};
    ws.getRow(1).font={bold:true,color:{argb:'FFFFFFFF'}};
    ws.getRow(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF2678B8'}};
    ws.getRow(1).alignment={horizontal:'center',vertical:'middle',wrapText:true};
    ws.getRow(1).height=42;
    [28,22,32,16,15,12,12,12,15,14,14,20,14,20,16,12,12,12,26,16,16,18,16,18,14,20,48].forEach((w,i)=>ws.getColumn(i+1).width=w);
    const buf=await wb.xlsx.writeBuffer(); saveAs(new Blob([buf]),title);
  };
  const exportSeparate = async () => {
    for(let i=0;i<files.length;i++) {
      const data=rows.filter(r=>r.source===files[i].name);
      if(data.length) await makeWorkbook(data,files[i].name.replace(/\.pdf$/i,'')+'.xlsx');
    }
  };
  return <main className="shell">
    <header className="top"><div className="brand"><div className="logo"><FileSpreadsheet size={24}/></div><div><h1>PDF to Excel</h1><p>Quotation data extractor</p></div></div><span className="pill">Manual conversion</span></header>
    <section className="hero"><div><p className="eyebrow">DOCUMENT WORKFLOW</p><h2>Turn PDF quotations into clean Excel sheets.</h2><p className="sub">Upload one or more quotations. Extract every line item into your fixed format, review uncertain values, then export separate Box App import templates or one combined template.</p></div><div className="heroIcon"><FileText size={54}/><span>.PDF</span></div></section>
    <section className="panel">
      <div className="sectionTitle"><div><h3>1. Select PDF files</h3><p>Text PDFs and scanned PDFs are supported through OCR.</p></div><span className="count">{files.length} selected</span></div>
      <label className="drop" onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();addFiles(e.dataTransfer.files)}}><input type="file" accept="application/pdf,.pdf" multiple onChange={e=>addFiles(e.target.files)}/><div className="uploadIcon"><Upload size={26}/></div><strong>Drop PDFs here or browse files</strong><span>Multiple PDFs supported · Files are processed in this browser</span></label>
      {files.length>0&&<div className="fileList">{files.map((f,i)=><div className="fileItem" key={i}><FileText size={18}/><span>{f.name}</span><small>{(f.size/1024).toFixed(0)} KB</small><button className="iconBtn" onClick={()=>setFiles(prev=>prev.filter((_,j)=>i!==j))} aria-label="Remove file"><Trash2 size={16}/></button></div>)}</div>}
      <button className="primary" disabled={!files.length||busy} onClick={convert}>{busy?<><LoaderCircle className="spin" size={18}/> Converting…</>:<><FileSpreadsheet size={18}/> Convert to Excel</>}</button>
      {status&&<p className="status">{status}</p>}
    </section>
    {rows.length>0&&<section className="panel">
      <div className="sectionTitle"><div><h3>2. Review extracted items</h3><p>Edit any cell that needs correction before exporting.</p></div><span className={reviewCount?'warn count':'ok count'}>{reviewCount?`${reviewCount} need review`:'All rows look complete'}</span></div>
      <div className="notice"><AlertTriangle size={17}/><span>Extraction is heuristic. Verify highlighted rows, item codes, quantities and dimensions against the source PDF before using the workbook.</span></div>
      <div className="tableWrap"><table><thead><tr>{IMPORT_HEADERS.map(h=><th key={h}>{h}</th>)}<th>CHECK</th></tr></thead><tbody>{rows.map((r,i)=><tr key={i} className={r.confidence==='review'?'needsReview':''}>{IMPORT_HEADERS.map(h=>{const vals=getTemplateValues(r); return <td key={h}><input value={String(vals[h] ?? '')} onChange={e=>updateRow(i,h,e.target.value)} aria-label={`${h} row ${i+1}`} /></td>})}<td>{r.confidence==='review'?<span className="reviewTag">Review</span>:<span className="goodTag">OK</span>}</td></tr>)}</tbody></table></div>
      <div className="exportRow"><button className="secondary" onClick={()=>makeWorkbook(rows,'Combined_Quotation_Items.xlsx')}><Download size={17}/> Export Box App template</button><button className="primary" onClick={exportSeparate}><Files size={17}/> Export separate Box templates</button></div>
    </section>}
    <footer>Exports the 27-column Box App Bulk Calculation Template with defaults filled. <span>For production use, verify every extracted value.</span></footer>
  </main>;
}

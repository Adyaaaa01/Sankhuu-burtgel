import React, { useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import mammoth from 'mammoth/mammoth.browser';
import * as pdfjsLib from 'pdfjs-dist';
import './styles.css';

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('pdfjs-dist/build/pdf.worker.mjs', import.meta.url).toString();

const money = (n) => Number(n || 0).toLocaleString('mn-MN');
const cleanNum = (v) => {
  if (v === null || v === undefined) return 0;
  if (typeof v === 'number') return v;
  const s = String(v).replace(/,/g, '').replace(/[₮\s]/g, '').replace(/[()]/g, '-');
  const m = s.match(/-?\d+(\.\d+)?/);
  return m ? Number(m[0]) : 0;
};
const isDateLike = (v) => /\d{4}[.\/-]\d{1,2}[.\/-]\d{1,2}|\d{1,2}[.\/-]\d{1,2}[.\/-]\d{4}/.test(String(v || ''));
const pick = (row, keys) => {
  const entries = Object.entries(row || {});
  for (const k of keys) {
    const found = entries.find(([h]) => String(h).toLowerCase().includes(k));
    if (found) return found[1];
  }
  return '';
};

async function readFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (['xlsx', 'xls'].includes(ext)) return readExcel(file);
  if (ext === 'csv') return readCsv(file);
  if (ext === 'pdf') return readPdf(file);
  if (ext === 'docx') return readDocx(file);
  if (ext === 'json') return readJson(file);
  return readText(file);
}

async function readExcel(file) {
  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
  const allRows = [];
  wb.SheetNames.forEach((sheetName) => {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
    rows.forEach((r) => allRows.push({ ...r, sourceSheet: sheetName }));
  });
  return rowsToTransactions(allRows, file.name);
}
async function readCsv(file) {
  const text = await file.text();
  const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
  return rowsToTransactions(parsed.data, file.name);
}
async function readJson(file) {
  const data = JSON.parse(await file.text());
  const rows = Array.isArray(data) ? data : Array.isArray(data.transactions) ? data.transactions : [data];
  return rowsToTransactions(rows, file.name);
}
async function readText(file) {
  return textToTransactions(await file.text(), file.name);
}
async function readDocx(file) {
  const buffer = await file.arrayBuffer();
  const result = await mammoth.extractRawText({ arrayBuffer: buffer });
  return textToTransactions(result.value, file.name);
}
async function readPdf(file) {
  const buffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  let text = '';
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    text += content.items.map((x) => x.str).join(' ') + '\n';
  }
  return textToTransactions(text, file.name);
}

function rowsToTransactions(rows, fileName) {
  return rows.map((row, i) => {
    const date = pick(row, ['огноо', 'date', 'posted', 'transaction date']) || Object.values(row).find(isDateLike) || '';
    const desc = pick(row, ['утга', 'гүйлгээ', 'description', 'memo', 'details', 'purpose']) || Object.values(row).slice(0, 5).join(' ');
    const income = cleanNum(pick(row, ['орлого', 'income', 'credit', 'deposit', 'inflow']));
    const expense = cleanNum(pick(row, ['зарлага', 'expense', 'debit', 'withdrawal', 'outflow']));
    const amount = income || expense || cleanNum(pick(row, ['amount', 'дүн', 'sum']));
    const type = income > 0 || (amount > 0 && !expense) ? 'Орлого' : 'Зарлага';
    return makeTransaction({ date, desc, amount: Math.abs(amount), type, source: fileName, raw: row }, i);
  }).filter((t) => t.amount > 0 || t.description.trim());
}

function textToTransactions(text, fileName) {
  const lines = text.split(/\n|\r/).map((x) => x.trim()).filter(Boolean);
  const tx = [];
  const lineRx = /(\d{4}[.\/-]\d{1,2}[.\/-]\d{1,2}|\d{1,2}[.\/-]\d{1,2}[.\/-]\d{4}).*?(-?\d[\d,\s]*\.?\d*)/g;
  lines.forEach((line, i) => {
    if (!isDateLike(line)) return;
    const nums = [...line.matchAll(/-?\d[\d,]*(?:\.\d+)?/g)].map((m) => cleanNum(m[0])).filter((n) => Math.abs(n) > 100);
    const amount = nums.length ? Math.abs(nums[nums.length - 1]) : 0;
    const type = /орлого|credit|deposit|income|cr\b/i.test(line) ? 'Орлого' : /зарлага|debit|withdraw|expense|dr\b|шимтгэл/i.test(line) ? 'Зарлага' : 'Зарлага';
    const date = (line.match(/\d{4}[.\/-]\d{1,2}[.\/-]\d{1,2}|\d{1,2}[.\/-]\d{1,2}[.\/-]\d{4}/) || [''])[0];
    tx.push(makeTransaction({ date, desc: line, amount, type, source: fileName, raw: line }, i));
  });
  if (!tx.length) {
    lines.slice(0, 100).forEach((line, i) => tx.push(makeTransaction({ date: '', desc: line, amount: 0, type: 'Зарлага', source: fileName, raw: line }, i)));
  }
  return tx;
}

function makeTransaction(input, i) {
  const lower = String(input.desc || '').toLowerCase();
  let code = '6000'; let account = 'Бусад зардал';
  if (input.type === 'Орлого') { code = '4000'; account = 'Борлуулалтын орлого'; }
  if (/шимтгэл|fee|charge|commission/.test(lower)) { code = '6200'; account = 'Банкны шимтгэл'; }
  if (/түрээс|rent/.test(lower)) { code = '6100'; account = 'Түрээсийн зардал'; }
  if (/цалин|salary|wage/.test(lower)) { code = '6300'; account = 'Цалингийн зардал'; }
  if (/ноат|нөат|vat/.test(lower)) { code = input.type === 'Орлого' ? '5100' : '5200'; account = input.type === 'Орлого' ? 'НӨАТ өглөг' : 'НӨАТ авлага'; }
  return { id: `${Date.now()}-${i}-${Math.random().toString(16).slice(2)}`, date: String(input.date || ''), description: String(input.desc || ''), type: input.type, amount: Number(input.amount || 0), code, account, source: input.source, raw: input.raw };
}

function makeJournal(transactions) {
  const rows = [];
  transactions.forEach((t) => {
    if (t.type === 'Орлого') {
      rows.push({ date: t.date, desc: t.description, account: 'Харилцах банк / Касс', code: '1100', debit: t.amount, credit: 0, source: t.source });
      rows.push({ date: t.date, desc: t.description, account: t.account, code: t.code, debit: 0, credit: t.amount, source: t.source });
    } else {
      rows.push({ date: t.date, desc: t.description, account: t.account, code: t.code, debit: t.amount, credit: 0, source: t.source });
      rows.push({ date: t.date, desc: t.description, account: 'Харилцах банк / Касс', code: '1100', debit: 0, credit: t.amount, source: t.source });
    }
  });
  return rows;
}
function groupByAccount(journal) {
  const map = {};
  journal.forEach((r) => {
    const key = `${r.code} ${r.account}`;
    if (!map[key]) map[key] = { account: key, debit: 0, credit: 0 };
    map[key].debit += r.debit; map[key].credit += r.credit;
  });
  return Object.values(map);
}

function App() {
  const [transactions, setTransactions] = useState([]);
  const [tab, setTab] = useState('dashboard');
  const [msg, setMsg] = useState('Файл оруулна уу. Excel, PDF, CSV, TXT, JSON, DOCX дэмжинэ.');
  const journal = useMemo(() => makeJournal(transactions), [transactions]);
  const accounts = useMemo(() => groupByAccount(journal), [journal]);
  const totals = useMemo(() => ({ income: transactions.filter(t => t.type === 'Орлого').reduce((a,b)=>a+b.amount,0), expense: transactions.filter(t => t.type === 'Зарлага').reduce((a,b)=>a+b.amount,0), debit: journal.reduce((a,b)=>a+b.debit,0), credit: journal.reduce((a,b)=>a+b.credit,0) }), [transactions, journal]);
  const profit = totals.income - totals.expense;

  async function onFiles(e) {
    const files = [...e.target.files];
    setMsg('Файлууд уншиж байна...');
    try {
      const batches = await Promise.all(files.map(readFile));
      const flat = batches.flat();
      setTransactions(flat);
      setMsg(`${files.length} файл уншлаа. ${flat.length} мөр илэрлээ. Дт ${money(makeJournal(flat).reduce((a,b)=>a+b.debit,0))} = Кт ${money(makeJournal(flat).reduce((a,b)=>a+b.credit,0))}`);
    } catch (err) { setMsg('Файл уншихад алдаа гарлаа: ' + err.message); }
  }
  function askAi() {
    const advice = [];
    if (totals.debit !== totals.credit) advice.push('Дт/Кт тэнцэхгүй байна. Journal кодлолтоо шалга.'); else advice.push('Дт/Кт тэнцэж байна.');
    if (totals.expense > totals.income * 0.7) advice.push('Зардал орлогын 70%-иас өндөр байна. Түрээс, шимтгэл, цалин, түүхий эдийг тусад нь бууруулах төлөвлөгөө гарга.');
    if (profit > 0) advice.push(`Ашигтай байна: ${money(profit)}₮. Борлуулалтын сувгаа нэмэх боломжтой.`); else advice.push(`Алдагдалтай байна: ${money(Math.abs(profit))}₮. Өдөр тутмын зарлагын лимит тогтоо.`);
    alert(advice.join('\n'));
  }
  const tabs = [['dashboard','Dashboard'], ['bank','Bank upload'], ['journal','Journal'], ['ozt','OZT'], ['tdans','T данс'], ['balance','Balance'], ['ai','AI зөвлөх']];
  return <div className="app">
    <header><h1>Vibe Cafe санхүү бүртгэл</h1><p>Бүх төрлийн банкны хуулга уншигч + журнал + OZT + T данс + баланс</p></header>
    <section className="upload"><input type="file" multiple accept=".xlsx,.xls,.csv,.pdf,.txt,.json,.docx" onChange={onFiles}/><button onClick={() => setTransactions([])}>Цэвэрлэх</button><span>{msg}</span></section>
    <nav>{tabs.map(([k,v]) => <button className={tab===k?'active':''} onClick={()=>setTab(k)} key={k}>{v}</button>)}</nav>
    {tab==='dashboard' && <Dashboard totals={totals} profit={profit} count={transactions.length}/>} 
    {tab==='bank' && <Table rows={transactions} cols={[['date','Огноо'],['description','Гүйлгээ'],['type','Төрөл'],['amount','Дүн'],['code','Код'],['account','Данс'],['source','Файл']]}/>} 
    {tab==='journal' && <Table rows={journal} cols={[['date','Огноо'],['desc','Гүйлгээ'],['code','Код'],['account','Данс'],['debit','Дт'],['credit','Кт'],['source','Файл']]}/>} 
    {tab==='ozt' && <Ozt accounts={accounts}/>} 
    {tab==='tdans' && <TDans accounts={accounts}/>} 
    {tab==='balance' && <Balance accounts={accounts} totals={totals} profit={profit}/>} 
    {tab==='ai' && <section className="card"><h2>AI зөвлөх</h2><p>Одоогоор browser дотор ажиллах demo зөвлөх. Дараа нь OpenAI API key холбовол жинхэнэ чат болно.</p><textarea placeholder="Жишээ: НӨАТ ба банкны орлого яагаад зөрөөд байна?"/><button onClick={askAi}>Зөвлөгөө авах</button></section>} 
  </div>;
}
function Dashboard({ totals, profit, count }) { return <section className="grid"><Card title="Орлого" value={`${money(totals.income)}₮`}/><Card title="Зарлага" value={`${money(totals.expense)}₮`}/><Card title="Ашиг / Алдагдал" value={`${money(profit)}₮`}/><Card title="Дт = Кт" value={`${money(totals.debit)} / ${money(totals.credit)}`}/><Card title="Гүйлгээ" value={count}/></section> }
function Card({title,value}) { return <div className="card"><h3>{title}</h3><strong>{value}</strong></div> }
function Table({ rows, cols }) { return <section className="card wide"><div className="table"><table><thead><tr>{cols.map(c=><th key={c[0]}>{c[1]}</th>)}</tr></thead><tbody>{rows.slice(0,1000).map((r,i)=><tr key={i}>{cols.map(c=><td key={c[0]}>{typeof r[c[0]]==='number'?money(r[c[0]]):String(r[c[0]]||'')}</td>)}</tr>)}</tbody></table></div></section> }
function Ozt({ accounts }) { return <section className="card wide"><h2>OZT Sheet</h2><Table rows={accounts.map(a=>({...a, balance:a.debit-a.credit}))} cols={[['account','Данс'],['debit','Дт'],['credit','Кт'],['balance','Үлдэгдэл']]}/></section> }
function TDans({ accounts }) { return <section className="tdans">{accounts.map((a,i)=><div className="t" key={i}><h3>{a.account}</h3><div><span>Дт {money(a.debit)}</span><span>Кт {money(a.credit)}</span></div></div>)}</section> }
function Balance({ accounts, totals, profit }) { const cash = accounts.find(a=>a.account.includes('1100')) || {debit:0,credit:0}; return <section className="card"><h2>Balance Sheet</h2><p>Мөнгөн хөрөнгө: <b>{money(cash.debit-cash.credit)}₮</b></p><p>Нийт орлого: <b>{money(totals.income)}₮</b></p><p>Нийт зарлага: <b>{money(totals.expense)}₮</b></p><p>Ашиг/алдагдал: <b>{money(profit)}₮</b></p><p>Шалгалт: <b>{totals.debit===totals.credit?'Тэнцсэн':'Тэнцээгүй'}</b></p></section> }

createRoot(document.getElementById('root')).render(<App />);

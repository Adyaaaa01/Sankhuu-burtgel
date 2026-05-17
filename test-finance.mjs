import XLSX from 'xlsx';
const path='../2025он-1-12р-сар-хуулга.xlsx';
const wb=XLSX.readFile(path);const ws=wb.Sheets['Bank'];
function dateFmt(v){if(!v)return'';if(typeof v==='number'){const d=XLSX.SSF.parse_date_code(v);return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`}return String(v).slice(0,10)}
function guessCode(desc='',inc=0,exp=0){let t=String(desc).toLowerCase();if(t.includes('шимтгэл')||t.includes('fee'))return 503;if(t.includes('түрээс'))return inc>0?404:502;if(t.includes('цахилгаан'))return 504;if(t.includes('ус')||t.includes('дулаан'))return 505;if(t.includes('карго'))return 506;if(exp>0)return 508;return 401}
const arr=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
const bank=arr.slice(1).filter(r=>r.some(Boolean)).map((r,i)=>({id:i+1,date:dateFmt(r[0]),desc:r[1]||'',income:+r[2]||0,expense:+r[3]||0,code:+r[4]||guessCode(r[1],+r[2]||0,+r[3]||0)}));
const journal=bank.flatMap(r=>{const amount=r.income||r.expense;if(!amount)return[];return r.income>0?[{debitCode:102,creditCode:r.code,debit:amount,credit:amount}]:[{debitCode:r.code,creditCode:102,debit:amount,credit:amount}]});
const balanced=journal.every(j=>Math.abs(j.debit-j.credit)<0.01);
console.log(JSON.stringify({bankRows:bank.length,journalRows:journal.length,balanced,totalIncome:bank.reduce((s,r)=>s+r.income,0),totalExpense:bank.reduce((s,r)=>s+r.expense,0)},null,2));
if(!balanced) process.exit(1);

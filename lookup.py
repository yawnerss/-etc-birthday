#!/usr/bin/env python3
import os
import re
import threading
import time
import unicodedata
from datetime import datetime, date
from difflib import SequenceMatcher

from flask import Flask, jsonify, request
from openpyxl import load_workbook

XLSX_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ofes_data.xlsx")
SHEET_NAME = "Sheet1"
FUZZY_MIN = 0.72

app = Flask(__name__)
_load_lock = threading.Lock()
_INDEX = {"key": None, "rows": [], "by_id": {}, "by_token": {}, "by_prefix": {}}

_BDAY_FORMATS = [
    "%B %d, %Y", "%b %d, %Y",
    "%B %d %Y", "%b %d %Y",
    "%d %B %Y", "%d %b %Y",
    "%Y-%m-%d", "%Y/%m/%d",
    "%m/%d/%Y", "%m/%d/%y",
    "%m-%d-%Y", "%m-%d-%y",
    "%d-%m-%Y", "%d-%m-%y",
]

_MONTH_NAMES = ["", "January", "February", "March", "April", "May", "June",
                "July", "August", "September", "October", "November", "December"]

_MONTH_RE = (r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|"
             r"Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|"
             r"Nov(?:ember)?|Dec(?:ember)?)")


def _try_parse_date(s):
    if not s:
        return None
    s = re.sub(r"\bSept\b", "Sep", s, flags=re.IGNORECASE)
    for fmt in _BDAY_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    for fmt in ("%B %d", "%b %d"):
        try:
            d = datetime.strptime(s, fmt)
            return date(1900, d.month, d.day)
        except ValueError:
            pass
    return None


def parse_birthday(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return _try_parse_date(str(v).strip())


def find_date_in_line(line):
    if not line:
        return None, None
    pats = [
        (rf"\b{_MONTH_RE}\s+\d{{1,2}}(?:,?\s+\d{{4}})?\b",
         ["%B %d, %Y", "%b %d, %Y", "%B %d %Y", "%b %d %Y", "%B %d", "%b %d"]),
        (rf"\b\d{{1,2}}\s+{_MONTH_RE}(?:,?\s+\d{{4}})?\b",
         ["%d %B %Y", "%d %b %Y", "%d %B", "%d %b"]),
        (r"\b\d{4}-\d{2}-\d{2}\b", ["%Y-%m-%d"]),
        (r"\b\d{4}/\d{2}/\d{2}\b", ["%Y/%m/%d"]),
        (r"\b\d{1,2}/\d{1,2}/\d{2,4}\b", ["%m/%d/%Y", "%m/%d/%y", "%d/%m/%Y", "%d/%m/%y"]),
        (r"\b\d{1,2}-\d{1,2}-\d{2,4}\b", ["%m-%d-%Y", "%m-%d-%y", "%d-%m-%Y", "%d-%m-%y"]),
    ]
    for pat, fmts in pats:
        m = re.search(pat, line, re.IGNORECASE)
        if not m:
            continue
        raw = m.group(0)
        for fmt in fmts:
            try:
                return datetime.strptime(raw, fmt).date(), raw
            except ValueError:
                try:
                    return datetime.strptime(raw.replace("Sept", "Sep"), fmt).date(), raw
                except ValueError:
                    pass
    return None, None


def _fold(s):
    s = unicodedata.normalize("NFKD", str(s or ""))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()


def _tokens(s):
    s = _fold(s)
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return [t for t in s.split() if t]


def name_match_score(query, candidate):
    q = _tokens(query)
    c = _tokens(candidate)
    if not q or not c:
        return 0.0, "none"
    if q == c:
        return 1.0, "exact"
    if sorted(q) == sorted(c) or set(q) == set(c):
        return 0.99, "tokens"
    worst = 1.0
    for qt in q:
        best = 0.0
        for ct in c:
            if qt == ct:
                best = 1.0
                break
            r = SequenceMatcher(None, qt, ct).ratio()
            if r > best:
                best = r
        worst = min(worst, best)
    full = SequenceMatcher(None, " ".join(sorted(q)), " ".join(sorted(c))).ratio()
    score = max(worst, full * 0.98)
    if set(q).issubset(set(c)):
        score = max(score, 0.93)
    if set(c).issubset(set(q)):
        score = max(score, 0.90)
    return min(score, 0.98), "fuzzy"


def parse_line_forms(line):
    s = (line or "").strip()
    if not s:
        return []
    forms = []
    if "\t" in s:
        parts = [p.strip() for p in s.split("\t") if p.strip()]
        if len(parts) >= 2:
            sur, first = parts[0], " ".join(parts[1:])
            forms.append(f"{sur} {first}")
            forms.append(f"{first} {sur}")
            forms.append(f"{sur}, {first}")
    if not forms and "," in s:
        parts = [p.strip() for p in s.split(",", 1)]
        if len(parts) == 2 and parts[0] and parts[1]:
            sur, first = parts
            forms.append(f"{sur} {first}")
            forms.append(f"{first} {sur}")
    if not forms and re.search(r"\s{2,}", s):
        parts = re.split(r"\s{2,}", s, maxsplit=1)
        if len(parts) == 2 and parts[0].strip() and parts[1].strip():
            sur, first = parts[0].strip(), parts[1].strip()
            forms.append(f"{sur} {first}")
            forms.append(f"{first} {sur}")
    forms.append(s)
    seen, out = set(), []
    for f in forms:
        k = re.sub(r"\s+", " ", _fold(f)).strip()
        if k and k not in seen:
            seen.add(k)
            out.append(f)
    return out


def _cell_to_str(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%B %d, %Y")
    if isinstance(v, date):
        return v.strftime("%B %d, %Y")
    return str(v)


def _load_index():
    if not os.path.exists(XLSX_PATH):
        raise FileNotFoundError(f"missing workbook: {XLSX_PATH}")
    wb = load_workbook(XLSX_PATH, read_only=True, data_only=True)
    try:
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb[wb.sheetnames[0]]
        it = ws.iter_rows(values_only=True)
        header_row = next(it, None)
        if header_row is None:
            return {"rows": [], "by_id": {}, "by_token": {}, "by_prefix": {}}

        headers = {}
        for idx, name in enumerate(header_row):
            if name is not None:
                headers[str(name).strip()] = idx

        def col(*cands):
            low = {k.lower(): v for k, v in headers.items()}
            for c in cands:
                if c.lower() in low:
                    return low[c.lower()]
            return None

        id_i = col("Student ID", "StudentID", "ID")
        name_i = col("Name")
        bd_i = col("Birthday")

        rows, by_id, by_token, by_prefix = [], {}, {}, {}
        for r_idx, row in enumerate(it, start=2):
            if row is None:
                continue
            sid = row[id_i] if id_i is not None and id_i < len(row) else None
            nm = row[name_i] if name_i is not None and name_i < len(row) else None
            bd = row[bd_i] if bd_i is not None and bd_i < len(row) else None
            sid_s = _cell_to_str(sid)
            nm_s = _cell_to_str(nm)
            if not sid_s and not nm_s:
                continue
            bd_s = _cell_to_str(bd)
            bd_parsed = parse_birthday(bd_s)
            name_tokens = _tokens(nm_s) if nm_s else []
            id_folded = re.sub(r"\s+", " ", _fold(sid_s)).strip() if sid_s else ""

            i = len(rows)
            rows.append({
                "row": r_idx,
                "id": sid_s,
                "name": nm_s,
                "birthday": bd_s,
                "birthday_month": bd_parsed.month if bd_parsed else None,
                "birthday_day": bd_parsed.day if bd_parsed else None,
                "name_tokens": tuple(name_tokens),
                "id_folded": id_folded,
            })
            if id_folded:
                by_id[id_folded] = i
            for t in name_tokens:
                by_token.setdefault(t, set()).add(i)
                if len(t) >= 3:
                    by_prefix.setdefault(t[:3], set()).add(i)

        return {"rows": rows, "by_id": by_id, "by_token": by_token, "by_prefix": by_prefix}
    finally:
        wb.close()


def get_index():
    global _INDEX
    try:
        st = os.stat(XLSX_PATH)
        key = (st.st_mtime_ns, st.st_size)
    except OSError as e:
        raise FileNotFoundError(f"missing workbook: {XLSX_PATH}") from e
    if _INDEX["key"] == key:
        return _INDEX
    with _load_lock:
        try:
            st = os.stat(XLSX_PATH)
            key = (st.st_mtime_ns, st.st_size)
        except OSError as e:
            raise FileNotFoundError(f"missing workbook: {XLSX_PATH}") from e
        if _INDEX["key"] != key:
            fresh = _load_index()
            fresh["key"] = key
            _INDEX = fresh
    return _INDEX


def match_index(idx, line, min_score=FUZZY_MIN):
    forms = parse_line_forms(line)
    if not forms:
        return None
    by_id = idx["by_id"]
    by_token = idx["by_token"]
    by_prefix = idx["by_prefix"]
    rows = idx["rows"]

    for form in forms:
        fk = re.sub(r"\s+", " ", _fold(form)).strip()
        if fk in by_id:
            return {"index": by_id[fk], "score": 1.0, "matchType": "id"}
    for form in forms:
        for tok in re.split(r"[,\s;]+", form):
            if not tok:
                continue
            tk = re.sub(r"\s+", " ", _fold(tok)).strip()
            if tk in by_id:
                return {"index": by_id[tk], "score": 1.0, "matchType": "id-token"}

    best = None
    for form in forms:
        qtok = set(_tokens(form))
        if not qtok:
            continue
        cands = set()
        for t in qtok:
            cands.update(by_token.get(t, ()))
            if len(t) >= 3:
                cands.update(by_prefix.get(t[:3], ()))
        if not cands:
            continue
        for i in cands:
            sc, mt = name_match_score(form, rows[i]["name"])
            if best is None or sc > best["score"]:
                best = {"index": i, "score": sc, "matchType": mt}
    if best and best["score"] >= min_score:
        return best
    return None


def student_from_index(idx, i):
    row = idx["rows"][i]
    return {"Student ID": row["id"], "Name": row["name"], "Birthday": row["birthday"]}


def resolve_month(raw):
    if raw is None or raw == "":
        d = datetime.now()
        return d.month, _MONTH_NAMES[d.month]
    s = str(raw).strip()
    if s.isdigit():
        n = int(s)
        if 1 <= n <= 12:
            return n, _MONTH_NAMES[n]
    low = s.lower()
    for i, name in enumerate(_MONTH_NAMES):
        if not name:
            continue
        if name.lower() == low or name[:3].lower() == low:
            return i, name
    return None, None


INDEX_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Birthdays</title>
<style>
:root{--bg:#fff;--panel:#fafafa;--bd:#e5e7eb;--fg:#111827;--muted:#6b7280;--ac:#2563eb;--ac-fg:#fff;--ok:#15803d;--bad:#b91c1c;--chip:#f3f4f6;}
html[data-theme="dark"]{--bg:#0b0d10;--panel:#121519;--bd:#22272e;--fg:#e6edf3;--muted:#8b949e;--ac:#58a6ff;--ac-fg:#0b0d10;--ok:#3fb950;--bad:#f85149;--chip:#1c2128;}
*{box-sizing:border-box;}
html,body{margin:0;background:var(--bg);color:var(--fg);}
body{font:14px/1.55 ui-sans-serif,-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;-webkit-font-smoothing:antialiased;}
header{display:flex;align-items:center;justify-content:space-between;padding:14px 20px;border-bottom:1px solid var(--bd);position:sticky;top:0;background:var(--bg);z-index:10;}
header h1{margin:0;font-size:14px;font-weight:600;letter-spacing:.01em;}
.toggle{border:1px solid var(--bd);border-radius:8px;background:var(--panel);color:var(--fg);padding:6px 10px;font:inherit;font-size:13px;cursor:pointer;display:inline-flex;gap:6px;align-items:center;}
.toggle:hover{border-color:var(--muted);}
main{max-width:1100px;margin:0 auto;padding:20px;}
section{margin-bottom:24px;}
h2{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);margin:0 0 10px;}
.field{display:flex;flex-direction:column;gap:4px;}
.field label{font-size:12px;color:var(--muted);}
input,select,textarea{width:100%;padding:8px 10px;border:1px solid var(--bd);border-radius:8px;background:var(--panel);color:var(--fg);font:inherit;}
input:focus,select:focus,textarea:focus{outline:2px solid var(--ac);outline-offset:-1px;}
textarea{min-height:140px;font-family:ui-monospace,SFMono-Regular,Menlo,monospace;font-size:13px;resize:vertical;white-space:pre;}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px;}
.row{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-top:12px;}
button{padding:8px 14px;border:1px solid var(--bd);border-radius:8px;background:var(--panel);color:var(--fg);font:inherit;font-size:13px;cursor:pointer;}
button.primary{background:var(--ac);border-color:var(--ac);color:var(--ac-fg);}
button:disabled{opacity:.5;cursor:not-allowed;}
#status{min-height:18px;font-size:13px;color:var(--muted);margin-bottom:10px;}
#status.ok{color:var(--ok);}
#status.bad{color:var(--bad);}
.spinner{display:inline-block;width:10px;height:10px;border:2px solid var(--muted);border-top-color:transparent;border-radius:50%;animation:spin .7s linear infinite;vertical-align:middle;margin-right:6px;}
@keyframes spin{to{transform:rotate(360deg);}}
.summary{display:flex;gap:16px;flex-wrap:wrap;margin-bottom:12px;font-size:13px;color:var(--muted);}
.summary b{color:var(--fg);font-weight:600;}
table{width:100%;border-collapse:collapse;font-size:13px;}
th,td{border-bottom:1px solid var(--bd);padding:8px 10px;text-align:left;vertical-align:top;}
th{font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);font-weight:600;}
td{word-break:break-word;}
tr:last-child td{border-bottom:none;}
.chip{display:inline-block;padding:1px 8px;border-radius:999px;background:var(--chip);color:var(--muted);font-size:11px;font-weight:500;}
.chip.on{background:color-mix(in srgb,var(--ok) 20%,transparent);color:var(--ok);}
.empty{color:var(--muted);font-style:italic;}
.group h3{font-size:12px;font-weight:600;color:var(--muted);margin:16px 0 6px;text-transform:uppercase;letter-spacing:.06em;}
</style>
</head>
<body>
<header>
  <h1>Birthdays</h1>
  <button class="toggle" id="theme"><span id="theme-ico">◐</span> <span id="theme-txt">Dark</span></button>
</header>
<main>
<section>
  <h2>Scan workbook</h2>
  <div class="grid">
    <div class="field"><label for="month">Month</label><select id="month"></select></div>
  </div>
  <div class="row"><button class="primary" id="scanBtn" onclick="scanWorkbook()">Scan</button></div>
</section>
<section>
  <h2>Classmate list</h2>
  <div class="field">
    <label for="paste">Paste names / IDs / txt lines (one per line — tab, comma, or multi-space separated)</label>
    <textarea id="paste" placeholder="BALINO&#9;HERIE-JAY&#10;BANGAYAN&#9;DOMINIC&#10;IGARTA&#9;ALVIN&#10;26-05959"></textarea>
  </div>
  <div class="row">
    <input type="file" id="file" accept=".txt,text/plain" style="max-width:240px">
    <button class="primary" id="checkBtn" onclick="checkList()">Check list</button>
    <button onclick="clearAll()">Clear</button>
  </div>
</section>
<section>
  <h2>Result</h2>
  <div id="status"></div>
  <div id="out"><p class="empty">No data loaded.</p></div>
</section>
</main>
<script>
const STATUS=document.getElementById('status');
const OUT=document.getElementById('out');
const MONTHS=["","January","February","March","April","May","June","July","August","September","October","November","December"];

function applyTheme(t){
  document.documentElement.setAttribute('data-theme',t);
  document.getElementById('theme-txt').textContent=t==='dark'?'Light':'Dark';
  document.getElementById('theme-ico').textContent=t==='dark'?'☀':'◐';
  localStorage.setItem('bday-theme',t);
}
(function(){
  const saved=localStorage.getItem('bday-theme');
  const sys=matchMedia('(prefers-color-scheme: dark)').matches?'dark':'light';
  applyTheme(saved||sys);
})();
document.getElementById('theme').addEventListener('click',()=>{
  applyTheme(document.documentElement.getAttribute('data-theme')==='dark'?'light':'dark');
});
(function(){
  const sel=document.getElementById('month');
  const cur=new Date().getMonth()+1;
  for(let i=1;i<=12;i++){
    const o=document.createElement('option');
    o.value=String(i);
    o.textContent=MONTHS[i]+(i===cur?' (current)':'');
    sel.appendChild(o);
  }
  sel.value=String(cur);
})();

function setStatus(m,k){
  STATUS.innerHTML='';
  if(k==='load'){STATUS.appendChild(Object.assign(document.createElement('span'),{className:'spinner'}));}
  STATUS.appendChild(document.createTextNode(m||''));
  STATUS.className=(k==='load')?'':(k||'');
}
function el(t,x){const e=document.createElement(t);if(x!==undefined)e.textContent=x;return e;}
function monthSel(){return parseInt(document.getElementById('month').value,10);}
function setBusy(b){
  document.getElementById('checkBtn').disabled=b;
  document.getElementById('scanBtn').disabled=b;
}

function renderTable(rows,cols){
  const t=el('table');
  const th=el('thead'),htr=el('tr');
  cols.forEach(c=>htr.appendChild(el('th',c)));
  th.appendChild(htr);t.appendChild(th);
  const tb=el('tbody');
  if(!rows||!rows.length){
    const tr=el('tr');const td=el('td','No entries.');td.colSpan=cols.length;td.className='empty';
    tr.appendChild(td);tb.appendChild(tr);
  }else{
    rows.forEach(r=>{
      const tr=el('tr');
      cols.forEach(c=>{
        const td=el('td');
        if(c==='This month'){
          const chip=el('span',r.thisMonth?'Yes':'No');
          chip.className='chip'+(r.thisMonth?' on':'');
          td.appendChild(chip);
        }else if(c==='Match'){
          if(r.matched){const chip=el('span',r.matchType+' · '+r.score);chip.className='chip';td.appendChild(chip);}
          else if(r.matchType==='inline'){const chip=el('span','inline date');chip.className='chip';td.appendChild(chip);}
          else{td.textContent='—';}
        }else{
          const src=r.student&&r.student[c]!==undefined?r.student[c]:r[c];
          td.textContent=src===undefined||src===null?'':String(src);
        }
        tr.appendChild(td);
      });
      tb.appendChild(tr);
    });
  }
  t.appendChild(tb);return t;
}

function renderWorkbookResult(data){
  OUT.innerHTML='';
  const s=el('div');s.className='summary';
  s.innerHTML='<span><b>'+data.count+'</b> born in <b>'+data.monthName+'</b></span>'+
    (data.elapsedMs!==undefined?'<span>'+data.elapsedMs+' ms</span>':'');
  OUT.appendChild(s);
  OUT.appendChild(renderTable(data.rows,['Student ID','Name','Birthday']));
}

function renderListResult(data){
  OUT.innerHTML='';
  const s=el('div');s.className='summary';
  s.innerHTML='<span><b>'+data.total+'</b> unique lines</span>'+
    (data.duplicates&&data.duplicates.length?'<span><b>'+data.duplicates.length+'</b> duplicate(s) skipped</span>':'')+
    '<span><b>'+data.matchedCount+'</b> matched</span>'+
    '<span><b>'+data.thisMonthCount+'</b> in <b>'+data.monthName+'</b></span>'+
    (data.unmatchedCount?'<span><b>'+data.unmatchedCount+'</b> unmatched</span>':'')+
    (data.elapsedMs!==undefined?'<span>'+data.elapsedMs+' ms</span>':'');
  OUT.appendChild(s);

  const g1=el('div');g1.className='group';
  g1.appendChild(el('h3','Birthday in '+data.monthName+' ('+data.thisMonth.length+')'));
  g1.appendChild(renderTable(data.thisMonth,['Student ID','Name','Birthday','This month','Match']));
  OUT.appendChild(g1);

  if(data.matchedOther.length){
    const g2=el('div');g2.className='group';
    g2.appendChild(el('h3','Other months ('+data.matchedOther.length+')'));
    g2.appendChild(renderTable(data.matchedOther,['Student ID','Name','Birthday','This month','Match']));
    OUT.appendChild(g2);
  }

  if(data.unmatched.length){
    const g3=el('div');g3.className='group';
    g3.appendChild(el('h3','Unmatched ('+data.unmatched.length+')'));
    const t=el('table');
    const th=el('thead'),htr=el('tr');
    ['Line','Reason'].forEach(c=>htr.appendChild(el('th',c)));
    th.appendChild(htr);t.appendChild(th);
    const tb=el('tbody');
    data.unmatched.forEach(u=>{
      const tr=el('tr');
      tr.appendChild(el('td',u.line));
      tr.appendChild(el('td',u.error||'no match'));
      tb.appendChild(tr);
    });
    t.appendChild(tb);g3.appendChild(t);OUT.appendChild(g3);
  }
}

async function withTimeout(url,opts,ms){
  const ctrl=new AbortController();
  const tm=setTimeout(()=>ctrl.abort(),ms);
  try{return await fetch(url,{...opts,signal:ctrl.signal});}
  finally{clearTimeout(tm);}
}

async function scanWorkbook(){
  setBusy(true);setStatus('Scanning…','load');
  try{
    const r=await withTimeout('/birthdays?month='+monthSel(),{},15000);
    const j=await r.json();
    if(!r.ok){setStatus(j.error||('Error '+r.status),'bad');return;}
    setStatus(j.count+' birthday(s) in '+j.monthName+' · '+j.elapsedMs+' ms','ok');
    renderWorkbookResult(j);
  }catch(e){
    setStatus(e.name==='AbortError'?'Timed out.':'Request failed: '+e.message,'bad');
  }finally{setBusy(false);}
}

async function checkList(){
  const text=document.getElementById('paste').value;
  const file=document.getElementById('file').files[0];
  if(!text.trim()&&!file){setStatus('Paste names or choose a txt file.','bad');return;}
  setBusy(true);setStatus('Checking…','load');
  try{
    let r;
    if(file){
      const fd=new FormData();
      fd.append('file',file);
      fd.append('month',String(monthSel()));
      r=await withTimeout('/birthdays/list',{method:'POST',body:fd},30000);
    }else{
      r=await withTimeout('/birthdays/list',{
        method:'POST',
        headers:{'Content-Type':'application/json'},
        body:JSON.stringify({text,month:monthSel()})
      },30000);
    }
    const j=await r.json();
    if(!r.ok){setStatus(j.error||('Error '+r.status),'bad');return;}
    setStatus(j.thisMonthCount+' classmate(s) have birthday in '+j.monthName+' · '+j.elapsedMs+' ms','ok');
    renderListResult(j);
  }catch(e){
    setStatus(e.name==='AbortError'?'Timed out.':'Request failed: '+e.message,'bad');
  }finally{setBusy(false);}
}

function clearAll(){
  document.getElementById('paste').value='';
  document.getElementById('file').value='';
  OUT.innerHTML='<p class="empty">No data loaded.</p>';
  setStatus('');
}
</script>
</body>
</html>
"""


@app.get("/")
def index():
    return INDEX_HTML, 200, {"Content-Type": "text/html; charset=utf-8"}


@app.get("/birthdays")
def birthdays():
    t0 = time.time()
    m_num, m_name = resolve_month(request.args.get("month"))
    if m_num is None:
        return jsonify({"error": "month must be 1-12 or month name"}), 400
    try:
        idx = get_index()
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 500
    out = []
    for row in idx["rows"]:
        if row["birthday_month"] != m_num:
            continue
        out.append({"Student ID": row["id"], "Name": row["name"], "Birthday": row["birthday"],
                    "_row": row["row"], "day": row["birthday_day"]})
    out.sort(key=lambda x: (x["day"] or 99, x.get("Name", "")))
    return jsonify({"month": m_num, "monthName": m_name, "count": len(out), "rows": out,
                    "elapsedMs": int((time.time() - t0) * 1000)}), 200


@app.post("/birthdays/list")
def birthdays_list():
    t0 = time.time()
    payload = request.get_json(silent=True)
    if payload is not None:
        text = payload.get("text", "")
        raw_month = payload.get("month")
    elif "file" in request.files:
        f = request.files["file"]
        text = f.read().decode("utf-8", errors="replace")
        raw_month = request.form.get("month")
    else:
        text = request.form.get("text", "")
        raw_month = request.form.get("month")

    raw_lines = [ln.strip() for ln in (text or "").splitlines() if ln.strip()]
    if not raw_lines:
        return jsonify({"error": "no lines provided"}), 400

    m_num, m_name = resolve_month(raw_month)
    if m_num is None:
        return jsonify({"error": "month must be 1-12 or month name"}), 400

    seen, lines, duplicates = set(), [], []
    for ln in raw_lines:
        k = re.sub(r"\s+", " ", _fold(ln)).strip()
        if k in seen:
            duplicates.append(ln)
            continue
        seen.add(k)
        lines.append(ln)

    try:
        idx = get_index()
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 500

    this_month, matched_other, unmatched = [], [], []

    for line in lines:
        inline_bd, matched_str = find_date_in_line(line)
        name_part = line.replace(matched_str, "").strip(" ,;-") if matched_str else line

        match = match_index(idx, line)
        if match is None and name_part and name_part != line:
            match = match_index(idx, name_part)

        entry = {"line": line}

        if match is not None:
            student = student_from_index(idx, match["index"])
            bd = parse_birthday(student["Birthday"]) or inline_bd
            entry.update({
                "matched": True,
                "rowNumber": student and idx["rows"][match["index"]]["row"],
                "score": round(match["score"], 3),
                "matchType": match["matchType"],
                "student": student,
            })
            if bd is not None:
                entry["birthdayMonth"] = bd.month
                entry["birthdayMonthName"] = _MONTH_NAMES[bd.month]
                entry["birthdayDay"] = bd.day
                entry["thisMonth"] = (bd.month == m_num)
            else:
                entry["birthdayMonth"] = None
                entry["birthdayMonthName"] = None
                entry["birthdayDay"] = None
                entry["thisMonth"] = False
        elif inline_bd is not None:
            entry.update({
                "matched": False, "rowNumber": None, "score": None, "matchType": "inline",
                "student": {"Student ID": "", "Name": name_part or line,
                            "Birthday": inline_bd.strftime("%B %d, %Y")},
                "birthdayMonth": inline_bd.month,
                "birthdayMonthName": _MONTH_NAMES[inline_bd.month],
                "birthdayDay": inline_bd.day,
                "thisMonth": inline_bd.month == m_num,
            })
        else:
            entry.update({
                "matched": False, "rowNumber": None, "score": None, "matchType": None,
                "student": None, "birthdayMonth": None, "birthdayMonthName": None,
                "birthdayDay": None, "thisMonth": False,
                "error": "no match and no inline date",
            })

        if entry["thisMonth"]:
            this_month.append(entry)
        elif entry["student"] is not None:
            matched_other.append(entry)
        else:
            unmatched.append(entry)

    this_month.sort(key=lambda e: (e.get("birthdayDay") or 99, (e.get("student") or {}).get("Name", "")))
    matched_other.sort(key=lambda e: (e.get("birthdayMonth") or 99, e.get("birthdayDay") or 99))

    return jsonify({
        "month": m_num, "monthName": m_name,
        "rawTotal": len(raw_lines), "total": len(lines), "duplicates": duplicates,
        "matchedCount": len(this_month) + len(matched_other),
        "thisMonthCount": len(this_month),
        "unmatchedCount": len(unmatched),
        "thisMonth": this_month, "matchedOther": matched_other, "unmatched": unmatched,
        "elapsedMs": int((time.time() - t0) * 1000),
    }), 200


@app.get("/health")
def health():
    return jsonify({"ok": True})


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
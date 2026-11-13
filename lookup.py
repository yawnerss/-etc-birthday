#!/usr/bin/env python3
import os
import re
import sys
import threading
import time
import traceback
import unicodedata
import urllib.request
from datetime import datetime, date
from difflib import SequenceMatcher

from flask import Flask, jsonify, request
from openpyxl import load_workbook

GITHUB_RAW = os.environ.get("XLSX_URL", "").strip()


def resolve_xlsx_path():
    here = os.path.dirname(os.path.abspath(__file__))

    env_path = (os.environ.get("XLSX_PATH") or "").strip()
    if env_path and os.path.exists(env_path):
        return env_path

    if GITHUB_RAW:
        tmp = "/tmp/ofes_data.xlsx"
        if os.path.exists(tmp):
            return tmp
    else:
        try:
            if os.path.exists("/tmp/ofes_data.xlsx"):
                os.remove("/tmp/ofes_data.xlsx")
        except OSError:
            pass

    for p in (os.path.join(here, "ofes_data.xlsx"),
              os.path.join(os.getcwd(), "ofes_data.xlsx")):
        if os.path.exists(p):
            return p

    return os.path.join(here, "ofes_data.xlsx")


XLSX_PATH = resolve_xlsx_path()
SHEET_NAME = "Sheet1"
FUZZY_MIN = 0.72

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024
_load_lock = threading.Lock()
_INDEX = {"key": None, "rows": [], "by_id": {}, "by_token": {}, "by_prefix": {},
          "by_sorted": {}, "sheet": SHEET_NAME}

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


def fetch_from_github():
    if not GITHUB_RAW:
        return None
    url = GITHUB_RAW
    if "?" not in url:
        url += "?t=" + str(int(time.time()))
    tmp = "/tmp/ofes_data.xlsx"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "lookup/1.0"})
        with urllib.request.urlopen(req, timeout=20) as r, open(tmp, "wb") as f:
            f.write(r.read())
        return tmp
    except Exception as e:
        print(f"[startup] github fetch failed: {type(e).__name__}: {e}", file=sys.stderr, flush=True)
        return None


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
        sheet = SHEET_NAME if SHEET_NAME in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet]
        it = ws.iter_rows(values_only=True)
        header_row = next(it, None)
        if header_row is None:
            return {"rows": [], "by_id": {}, "by_token": {}, "by_prefix": {},
                    "by_sorted": {}, "sheet": sheet}

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

        if name_i is None:
            raise ValueError(
                f"header 'Name' not found. Got headers: {list(headers.keys())}"
            )

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

        by_sorted = {}
        for i, r in enumerate(rows):
            if not r["name_tokens"]:
                continue
            key = tuple(sorted(r["name_tokens"]))
            by_sorted.setdefault(key, []).append(i)

        return {"rows": rows, "by_id": by_id, "by_token": by_token,
                "by_prefix": by_prefix, "by_sorted": by_sorted, "sheet": sheet}
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
    by_sorted = idx["by_sorted"]
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

    for form in forms:
        qtok = tuple(sorted(_tokens(form)))
        if not qtok:
            continue
        hit = by_sorted.get(qtok)
        if hit:
            return {"index": hit[0], "score": 0.99, "matchType": "tokens"}

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


def full_student(idx, i):
    row = idx["rows"][i]
    m_name = _MONTH_NAMES[row["birthday_month"]] if row["birthday_month"] else ""
    return {
        "Student ID": row["id"],
        "Name": row["name"],
        "Birthday": row["birthday"],
        "Month": m_name,
        "Day": row["birthday_day"] if row["birthday_day"] is not None else "",
        "_row": row["row"],
    }


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


def _err(where, e, t0):
    print(f"[{where}] {type(e).__name__}: {e}", file=sys.stderr, flush=True)
    traceback.print_exc(file=sys.stderr)
    sys.stderr.flush()
    return jsonify({
        "error": f"{type(e).__name__}: {e}",
        "where": where,
        "xlsx_path": XLSX_PATH,
        "exists": os.path.exists(XLSX_PATH),
        "elapsedMs": int((time.time() - t0) * 1000),
    }), 500


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
header{display:flex;align-items:center;justify-content:space-between;padding:14px 20px;border-bottom:1px solid var(--bd);position:sticky;top:0;background:var(--bg);z-index:10;gap:12px;}
header h1{margin:0;font-size:14px;font-weight:600;letter-spacing:.01em;}
.hbtns{display:flex;align-items:center;gap:8px;}
#serverStatus{font-size:11px;color:var(--muted);font-family:ui-monospace,SFMono-Regular,Menlo,monospace;}
.toggle{border:1px solid var(--bd);border-radius:8px;background:var(--panel);color:var(--fg);padding:6px 10px;font:inherit;font-size:13px;cursor:pointer;display:inline-flex;gap:6px;align-items:center;}
.toggle:hover{border-color:var(--muted);}
main{max-width:1200px;margin:0 auto;padding:20px;}
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
button:hover{filter:brightness(1.04);}
button:disabled{opacity:.5;cursor:not-allowed;}
#status{min-height:18px;font-size:13px;color:var(--muted);margin-bottom:10px;white-space:pre-wrap;word-break:break-word;}
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
.card{border:1px solid var(--bd);border-radius:10px;background:var(--panel);padding:16px;}
.card .kv{display:grid;grid-template-columns:140px 1fr;gap:6px 12px;font-size:13px;}
.card .kv .k{color:var(--muted);font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.05em;}
</style>
</head>
<body>
<header>
  <h1>Birthdays</h1>
  <div class="hbtns">
    <span id="serverStatus">●</span>
    <button class="toggle" id="theme"><span id="theme-ico">◐</span> <span id="theme-txt">Dark</span></button>
  </div>
</header>
<main>
<section>
  <h2>Single student lookup</h2>
  <div class="grid">
    <div class="field">
      <label for="single">Name or Student ID</label>
      <input id="single" placeholder="IGARTA ALVIN · ALVIN IGARTA · 26-05959" autocomplete="off">
    </div>
  </div>
  <div class="row"><button class="primary" id="singleBtn" onclick="lookupSingle()">Look up</button></div>
</section>
<section>
  <h2>Scan workbook</h2>
  <div class="grid">
    <div class="field"><label for="month">Month</label><select id="month"></select></div>
  </div>
  <div class="row">
    <button class="primary" id="scanBtn" onclick="scanWorkbook()">Scan month</button>
    <button class="primary" id="allBtn" onclick="scanAll()">Scan whole file</button>
  </div>
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
const PASTE=document.getElementById('paste');
const SINGLE=document.getElementById('single');
const SERVER_STATUS=document.getElementById('serverStatus');
const MONTHS=["","January","February","March","April","May","June","July","August","September","October","November","December"];

const BATCH_SIZE=400;
const RENDER_CAP=300;
const LS_PASTE='bday-paste';

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
(function(){
  const saved=localStorage.getItem(LS_PASTE);
  if(saved) PASTE.value=saved;
})();
let _t=null;
PASTE.addEventListener('input',()=>{
  clearTimeout(_t);
  _t=setTimeout(()=>localStorage.setItem(LS_PASTE,PASTE.value),400);
});

async function checkServerStatus(){
  SERVER_STATUS.style.color='var(--muted)';
  SERVER_STATUS.textContent='● checking';
  try{
    const r=await fetch('/healthz',{cache:'no-store'});
    const j=await r.json();
    if(j.ok){
      SERVER_STATUS.style.color='var(--ok)';
      SERVER_STATUS.textContent='● '+j.rows+' rows';
      SERVER_STATUS.title=j.path||'';
    }else{
      SERVER_STATUS.style.color='var(--bad)';
      SERVER_STATUS.textContent='● '+((j.error||'error').slice(0,60));
      SERVER_STATUS.title=(j.error||'')+'\n'+(j.path||'');
    }
  }catch(e){
    SERVER_STATUS.style.color='var(--bad)';
    SERVER_STATUS.textContent='● offline';
    SERVER_STATUS.title=String(e);
  }
}
checkServerStatus();

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
  document.getElementById('allBtn').disabled=b;
  document.getElementById('singleBtn').disabled=b;
}

function renderTable(rows,cols,cap){
  const t=el('table');
  const th=el('thead'),htr=el('tr');
  cols.forEach(c=>htr.appendChild(el('th',c)));
  th.appendChild(htr);t.appendChild(th);
  const tb=el('tbody');
  const total=rows?rows.length:0;
  if(!total){
    const tr=el('tr');const td=el('td','No entries.');td.colSpan=cols.length;td.className='empty';
    tr.appendChild(td);tb.appendChild(tr);
  }else{
    const shown=cap&&total>cap?rows.slice(0,cap):rows;
    shown.forEach(r=>{
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
    if(cap&&total>cap){
      const tr=el('tr');const td=el('td','Showing '+cap+' of '+total+' — download CSV for the full result.');
      td.colSpan=cols.length;td.className='empty';
      tr.appendChild(td);tb.appendChild(tr);
    }
  }
  t.appendChild(tb);return t;
}

function csvDownload(rows,name){
  const csv=rows.map(r=>r.map(v=>'"'+String(v==null?'':v).replace(/"/g,'""')+'"').join(',')).join('\n');
  const blob=new Blob([csv],{type:'text/csv;charset=utf-8'});
  const a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  a.download=name;
  a.click();
  setTimeout(()=>URL.revokeObjectURL(a.href),1000);
}

function renderSingle(data){
  OUT.innerHTML='';
  const s=el('div');s.className='summary';
  s.innerHTML='<span>matched <b>'+data.matchType+'</b> · score <b>'+data.score+'</b></span>'+
    '<span>row <b>'+data.student._row+'</b></span>'+
    '<span>query <b>'+data.query+'</b></span>'+
    (data.elapsedMs!==undefined?'<span>'+data.elapsedMs+' ms</span>':'');
  OUT.appendChild(s);

  const card=el('div');card.className='card';
  const kv=el('div');kv.className='kv';
  const rows=[
    ['Student ID', data.student['Student ID']],
    ['Name', data.student['Name']],
    ['Birthday', data.student['Birthday']],
    ['Month', data.student['Month']],
    ['Day', data.student['Day']],
    ['Row', data.student._row],
  ];
  rows.forEach(([k,v])=>{
    kv.appendChild(el('div',k)).className='k';
    kv.appendChild(el('div',v===''||v==null?'—':String(v)));
  });
  card.appendChild(kv);
  OUT.appendChild(card);

  const bar=el('div');bar.className='row';
  const btn=el('button','Copy JSON');
  btn.onclick=()=>navigator.clipboard.writeText(JSON.stringify(data.student,null,2));
  bar.appendChild(btn);
  OUT.appendChild(bar);
}

function renderWorkbookResult(data){
  OUT.innerHTML='';
  const s=el('div');s.className='summary';
  s.innerHTML='<span><b>'+data.count+'</b> born in <b>'+data.monthName+'</b></span>'+
    (data.elapsedMs!==undefined?'<span>'+data.elapsedMs+' ms</span>':'');
  OUT.appendChild(s);
  OUT.appendChild(renderTable(data.rows,['Student ID','Name','Birthday'],RENDER_CAP));
}

function renderAllResult(data){
  OUT.innerHTML='';
  const s=el('div');s.className='summary';
  s.innerHTML='<span><b>'+data.count+'</b> total row(s)</span>'+
    '<span>sheet <b>'+data.sheet+'</b></span>'+
    (data.elapsedMs!==undefined?'<span>'+data.elapsedMs+' ms</span>':'');
  OUT.appendChild(s);

  const bar=el('div');bar.className='row';bar.style.marginTop='0';
  const btn=el('button','Download CSV ('+data.count+' rows)');
  btn.onclick=()=>{
    const rows=[['Student ID','Name','Birthday','Month','Day']];
    data.rows.forEach(r=>rows.push([r['Student ID'],r['Name'],r['Birthday'],r['Month'],r['Day']]));
    csvDownload(rows,'ofes-all.csv');
  };
  bar.appendChild(btn);
  OUT.appendChild(bar);

  const bar2=el('div');bar2.className='row';bar2.style.marginTop='8px';
  const filt=el('input');
  filt.placeholder='Filter by name / ID / month…';
  filt.style.maxWidth='340px';
  bar2.appendChild(filt);
  OUT.appendChild(bar2);

  const holder=el('div');
  OUT.appendChild(holder);

  const draw=(q)=>{
    const qq=(q||'').trim().toLowerCase();
    const filtered=qq
      ? data.rows.filter(r=>
          String(r['Student ID']).toLowerCase().includes(qq)||
          String(r['Name']).toLowerCase().includes(qq)||
          String(r['Month']).toLowerCase().includes(qq)||
          String(r['Birthday']).toLowerCase().includes(qq))
      : data.rows;
    holder.innerHTML='';
    const sub=el('div');sub.className='summary';
    sub.innerHTML='<span>showing <b>'+Math.min(filtered.length,RENDER_CAP)+'</b> of <b>'+filtered.length+'</b></span>';
    holder.appendChild(sub);
    holder.appendChild(renderTable(filtered,['Student ID','Name','Birthday','Month','Day'],RENDER_CAP));
  };
  filt.addEventListener('input',()=>draw(filt.value));
  draw('');
}

function downloadAggCSV(agg){
  if(!agg) return;
  const rows=[['Group','Student ID','Name','Birthday','MatchType','Score']];
  const push=(g,e)=>{
    const s=e.student||{};
    rows.push([g,s['Student ID']||'',s['Name']||e.line||'',s['Birthday']||'',e.matchType||'',e.score||'']);
  };
  agg.thisMonth.forEach(e=>push('thisMonth',e));
  agg.matchedOther.forEach(e=>push('other',e));
  agg.unmatched.forEach(e=>rows.push(['unmatched','',''+(e.line||''),'','','']));
  csvDownload(rows,'birthdays-'+agg.monthName+'.csv');
}

function renderListResult(agg){
  OUT.innerHTML='';
  const s=el('div');s.className='summary';
  s.innerHTML='<span><b>'+agg.total+'</b> unique lines</span>'+
    (agg.duplicates&&agg.duplicates.length?'<span><b>'+agg.duplicates.length+'</b> duplicate(s) skipped</span>':'')+
    '<span><b>'+agg.matchedCount+'</b> matched</span>'+
    '<span><b>'+agg.thisMonthCount+'</b> in <b>'+(agg.monthName||'—')+'</b></span>'+
    (agg.unmatchedCount?'<span><b>'+agg.unmatchedCount+'</b> unmatched</span>':'')+
    (agg.failedBatches?'<span class="bad"><b>'+agg.failedBatches+'</b> batch(es) failed</span>':'')+
    (agg.elapsedMs!==undefined?'<span>'+agg.elapsedMs+' ms</span>':'');
  OUT.appendChild(s);

  if(agg.lastError){
    const e=el('div');e.className='card';
    e.style.borderColor='var(--bad)';
    const kv=el('div');kv.className='kv';
    kv.appendChild(el('div','Server error')).className='k';
    kv.appendChild(el('div',agg.lastError));
    e.appendChild(kv);
    OUT.appendChild(e);
    OUT.appendChild(el('div')).style.height='12px';
  }

  if(!agg.matchedCount && !agg.unmatchedCount){
    return;
  }

  const bar=el('div');bar.className='row';bar.style.marginTop='0';
  const btn=el('button','Download CSV (all '+agg.total+' lines)');
  btn.onclick=()=>downloadAggCSV(agg);
  bar.appendChild(btn);
  OUT.appendChild(bar);

  const g1=el('div');g1.className='group';
  g1.appendChild(el('h3','Birthday in '+(agg.monthName||'—')+' ('+agg.thisMonth.length+')'));
  g1.appendChild(renderTable(agg.thisMonth,['Student ID','Name','Birthday','This month','Match'],RENDER_CAP));
  OUT.appendChild(g1);

  if(agg.matchedOther.length){
    const g2=el('div');g2.className='group';
    g2.appendChild(el('h3','Other months ('+agg.matchedOther.length+')'));
    g2.appendChild(renderTable(agg.matchedOther,['Student ID','Name','Birthday','This month','Match'],RENDER_CAP));
    OUT.appendChild(g2);
  }

  if(agg.unmatched.length){
    const g3=el('div');g3.className='group';
    g3.appendChild(el('h3','Unmatched ('+agg.unmatched.length+')'));
    const t=el('table');
    const th=el('thead'),htr=el('tr');
    ['Line','Reason'].forEach(c=>htr.appendChild(el('th',c)));
    th.appendChild(htr);t.appendChild(th);
    const tb=el('tbody');
    const shown=agg.unmatched.length>RENDER_CAP?agg.unmatched.slice(0,RENDER_CAP):agg.unmatched;
    shown.forEach(u=>{
      const tr=el('tr');
      tr.appendChild(el('td',u.line));
      tr.appendChild(el('td',u.error||'no match'));
      tb.appendChild(tr);
    });
    if(agg.unmatched.length>RENDER_CAP){
      const tr=el('tr');const td=el('td','Showing '+RENDER_CAP+' of '+agg.unmatched.length+'.');
      td.colSpan=2;td.className='empty';tr.appendChild(td);tb.appendChild(tr);
    }
    t.appendChild(tb);g3.appendChild(t);OUT.appendChild(g3);
  }
}

async function withTimeout(url,opts,ms){
  const ctrl=new AbortController();
  const tm=setTimeout(()=>ctrl.abort(),ms);
  try{return await fetch(url,{...opts,signal:ctrl.signal});}
  finally{clearTimeout(tm);}
}

async function readErr(r){
  try{const j=await r.json();return j.error||('HTTP '+r.status);}
  catch(_){return 'HTTP '+r.status;}
}

async function lookupSingle(){
  const q=SINGLE.value.trim();
  if(!q){setStatus('Enter a name or student ID.','bad');return;}
  setBusy(true);setStatus('Looking up…','load');
  try{
    const r=await withTimeout('/student?q='+encodeURIComponent(q),{},15000);
    if(!r.ok){
      const e=await readErr(r);
      setStatus('Error: '+e,'bad');
      return;
    }
    const j=await r.json();
    if(!j.found){setStatus('No match for "'+q+'".','bad');return;}
    setStatus('Matched '+j.student['Name']+' · '+j.matchType+' '+j.score+' · '+j.elapsedMs+' ms','ok');
    renderSingle(j);
  }catch(e){
    setStatus(e.name==='AbortError'?'Timed out.':'Request failed: '+e.message,'bad');
  }finally{setBusy(false);checkServerStatus();}
}

async function scanWorkbook(){
  setBusy(true);setStatus('Scanning month…','load');
  try{
    const r=await withTimeout('/birthdays?month='+monthSel(),{},15000);
    if(!r.ok){
      const e=await readErr(r);
      setStatus('Error: '+e,'bad');
      return;
    }
    const j=await r.json();
    setStatus(j.count+' birthday(s) in '+j.monthName+' · '+j.elapsedMs+' ms','ok');
    renderWorkbookResult(j);
  }catch(e){
    setStatus(e.name==='AbortError'?'Timed out.':'Request failed: '+e.message,'bad');
  }finally{setBusy(false);checkServerStatus();}
}

async function scanAll(){
  setBusy(true);setStatus('Scanning whole file…','load');
  try{
    const r=await withTimeout('/all',{},30000);
    if(!r.ok){
      const e=await readErr(r);
      setStatus('Error: '+e,'bad');
      return;
    }
    const j=await r.json();
    setStatus(j.count+' row(s) in '+j.sheet+' · '+j.elapsedMs+' ms','ok');
    renderAllResult(j);
  }catch(e){
    setStatus(e.name==='AbortError'?'Timed out.':'Request failed: '+e.message,'bad');
  }finally{setBusy(false);checkServerStatus();}
}

async function checkList(){
  const raw=PASTE.value;
  const file=document.getElementById('file').files[0];
  if(!raw.trim()&&!file){setStatus('Paste names or choose a txt file.','bad');return;}

  const month=monthSel();
  const t0=performance.now();
  setBusy(true);

  if(file){
    setStatus('Uploading…','load');
    try{
      const fd=new FormData();
      fd.append('file',file);
      fd.append('month',String(month));
      const r=await withTimeout('/birthdays/list',{method:'POST',body:fd},120000);
      if(!r.ok){
        const e=await readErr(r);
        setStatus('Error: '+e,'bad');
        return;
      }
      const j=await r.json();
      j.elapsedMs=Math.round(performance.now()-t0);
      setStatus(j.thisMonthCount+' classmate(s) have birthday in '+j.monthName+' · '+j.elapsedMs+' ms','ok');
      renderListResult(j);
    }catch(e){
      setStatus(e.name==='AbortError'?'Timed out.':'Request failed: '+e.message,'bad');
    }finally{setBusy(false);checkServerStatus();}
    return;
  }

  const lines=raw.split('\n').map(s=>s.trim()).filter(Boolean);
  const batches=[];
  for(let i=0;i<lines.length;i+=BATCH_SIZE) batches.push(lines.slice(i,i+BATCH_SIZE));

  const agg={
    monthName:null, total:0, rawTotal:0, duplicates:[],
    thisMonth:[], matchedOther:[], unmatched:[],
    failedBatches:0, lastError:null
  };

  try{
    for(let b=0;b<batches.length;b++){
      setStatus('Batch '+(b+1)+'/'+batches.length+' ('+batches[b].length+' lines)…','load');
      try{
        const r=await withTimeout('/birthdays/list',{
          method:'POST',
          headers:{'Content-Type':'application/json'},
          body:JSON.stringify({text:batches[b].join('\n'),month})
        },60000);
        if(!r.ok){
          agg.failedBatches++;
          if(!agg.lastError) agg.lastError=await readErr(r);
          continue;
        }
        let j;
        try{ j=await r.json(); }
        catch(err){ agg.failedBatches++; if(!agg.lastError) agg.lastError='bad JSON from server (HTTP '+r.status+')'; continue; }
        agg.monthName=j.monthName;
        agg.rawTotal+=j.rawTotal||0;
        agg.total+=j.total||0;
        agg.duplicates.push(...(j.duplicates||[]));
        agg.thisMonth.push(...(j.thisMonth||[]));
        agg.matchedOther.push(...(j.matchedOther||[]));
        agg.unmatched.push(...(j.unmatched||[]));
      }catch(e){
        agg.failedBatches++;
        if(!agg.lastError) agg.lastError=String(e.message||e);
      }
    }

    agg.thisMonth.sort((a,b)=>((a.birthdayDay||99)-(b.birthdayDay||99))||String((a.student||{}).Name||'').localeCompare(String((b.student||{}).Name||'')));
    agg.matchedOther.sort((a,b)=>((a.birthdayMonth||99)-(b.birthdayMonth||99))||((a.birthdayDay||99)-(b.birthdayDay||99)));
    agg.matchedCount=agg.thisMonth.length+agg.matchedOther.length;
    agg.thisMonthCount=agg.thisMonth.length;
    agg.unmatchedCount=agg.unmatched.length;
    agg.elapsedMs=Math.round(performance.now()-t0);

    const tag=agg.failedBatches?(' · '+agg.failedBatches+' batch(es) failed: '+(agg.lastError||'')):'';
    setStatus(agg.thisMonthCount+' classmate(s) have birthday in '+(agg.monthName||'—')+' · '+agg.total+' lines · '+agg.elapsedMs+' ms'+tag,agg.failedBatches?'bad':'ok');
    renderListResult(agg);
  }finally{
    setBusy(false);checkServerStatus();
  }
}

function clearAll(){
  PASTE.value='';
  localStorage.removeItem(LS_PASTE);
  document.getElementById('file').value='';
  SINGLE.value='';
  OUT.innerHTML='<p class="empty">No data loaded.</p>';
  setStatus('');
}

SINGLE.addEventListener('keydown',e=>{if(e.key==='Enter')lookupSingle();});
</script>
</body>
</html>
"""


@app.get("/")
def index():
    return INDEX_HTML, 200, {"Content-Type": "text/html; charset=utf-8"}


@app.get("/debug/status")
def debug_status():
    out = {
        "xlsx_path": XLSX_PATH,
        "exists": os.path.exists(XLSX_PATH),
        "github_url": GITHUB_RAW or None,
        "sheet_name_setting": SHEET_NAME,
        "cwd": os.getcwd(),
        "script_dir": os.path.dirname(os.path.abspath(__file__)),
    }
    if out["exists"]:
        try:
            st = os.stat(XLSX_PATH)
            out["size"] = st.st_size
            out["mtime"] = datetime.fromtimestamp(st.st_mtime).isoformat()
            with open(XLSX_PATH, "rb") as f:
                head = f.read(16)
            out["first_bytes_hex"] = head.hex()
            out["looks_like_xlsx"] = head[:2] == b"PK"
        except OSError as e:
            out["stat_error"] = f"{type(e).__name__}: {e}"
        try:
            wb = load_workbook(XLSX_PATH, read_only=True, data_only=True)
            try:
                out["sheets"] = wb.sheetnames
                sname = SHEET_NAME if SHEET_NAME in wb.sheetnames else wb.sheetnames[0]
                out["active_sheet"] = sname
                ws = wb[sname]
                it = ws.iter_rows(values_only=True)
                header = next(it, None)
                out["headers"] = [str(h) if h is not None else None for h in (header or [])]
                cnt = 0
                for _ in it:
                    cnt += 1
                out["data_rows"] = cnt
            finally:
                wb.close()
        except Exception as e:
            out["load_error"] = f"{type(e).__name__}: {e}"
    try:
        idx = get_index()
        out["index_rows"] = len(idx["rows"])
        out["index_sheet"] = idx.get("sheet", SHEET_NAME)
    except Exception as e:
        out["index_error"] = f"{type(e).__name__}: {e}"
    return jsonify(out), 200


@app.get("/student")
def student_single():
    t0 = time.time()
    try:
        q = (request.args.get("q") or "").strip()
        if not q:
            return jsonify({"error": "q required"}), 400
        idx = get_index()
        match = match_index(idx, q)
        if match is None:
            return jsonify({"found": False, "query": q,
                            "elapsedMs": int((time.time() - t0) * 1000)}), 404
        return jsonify({
            "found": True, "query": q,
            "matchType": match["matchType"],
            "score": round(match["score"], 3),
            "student": full_student(idx, match["index"]),
            "elapsedMs": int((time.time() - t0) * 1000),
        }), 200
    except Exception as e:
        return _err("/student", e, t0)


@app.get("/birthdays")
def birthdays():
    t0 = time.time()
    try:
        m_num, m_name = resolve_month(request.args.get("month"))
        if m_num is None:
            return jsonify({"error": "month must be 1-12 or month name"}), 400
        idx = get_index()
        out = []
        for row in idx["rows"]:
            if row["birthday_month"] != m_num:
                continue
            out.append({"Student ID": row["id"], "Name": row["name"],
                        "Birthday": row["birthday"], "_row": row["row"],
                        "day": row["birthday_day"]})
        out.sort(key=lambda x: (x["day"] or 99, x.get("Name", "")))
        return jsonify({"month": m_num, "monthName": m_name, "count": len(out),
                        "rows": out, "elapsedMs": int((time.time() - t0) * 1000)}), 200
    except Exception as e:
        return _err("/birthdays", e, t0)


@app.get("/all")
def all_rows():
    t0 = time.time()
    try:
        idx = get_index()
        out = []
        for row in idx["rows"]:
            out.append({
                "Student ID": row["id"], "Name": row["name"],
                "Birthday": row["birthday"],
                "Month": _MONTH_NAMES[row["birthday_month"]] if row["birthday_month"] else "",
                "Day": row["birthday_day"] if row["birthday_day"] is not None else "",
                "_row": row["row"],
            })
        q = (request.args.get("q") or "").strip().lower()
        if q:
            out = [r for r in out if (
                q in str(r["Student ID"]).lower()
                or q in str(r["Name"]).lower()
                or q in str(r["Birthday"]).lower()
                or q in str(r["Month"]).lower()
            )]
        m_raw = request.args.get("month")
        if m_raw:
            m_num, m_name = resolve_month(m_raw)
            if m_num is None:
                return jsonify({"error": "month must be 1-12 or month name"}), 400
            out = [r for r in out if r["Month"] == m_name]
        return jsonify({"sheet": idx.get("sheet", SHEET_NAME), "count": len(out),
                        "rows": out, "elapsedMs": int((time.time() - t0) * 1000)}), 200
    except Exception as e:
        return _err("/all", e, t0)


@app.post("/birthdays/list")
def birthdays_list():
    t0 = time.time()
    try:
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

        idx = get_index()

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
                    "rowNumber": idx["rows"][match["index"]]["row"],
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
    except Exception as e:
        return _err("/birthdays/list", e, t0)


@app.get("/debug/match")
def debug_match():
    q = (request.args.get("q") or "").strip()
    if not q:
        return jsonify({"error": "q required"}), 400
    try:
        idx = get_index()
    except Exception as e:
        return _err("/debug/match", e, time.time())
    forms = parse_line_forms(q)
    by_token = idx["by_token"]
    by_prefix = idx["by_prefix"]
    by_sorted = idx["by_sorted"]
    detail = []
    for form in forms:
        ftok = _tokens(form)
        ftok_sorted = tuple(sorted(ftok))
        cands = set()
        for t in set(ftok):
            cands.update(by_token.get(t, ()))
            if len(t) >= 3:
                cands.update(by_prefix.get(t[:3], ()))
        scored = []
        for i in list(cands)[:20]:
            sc, mt = name_match_score(form, idx["rows"][i]["name"])
            scored.append({"index": i, "row": idx["rows"][i]["row"],
                           "name": idx["rows"][i]["name"],
                           "score": round(sc, 3), "matchType": mt})
        scored.sort(key=lambda x: -x["score"])
        detail.append({"form": form, "tokens": ftok,
                       "sorted_key_hit": bool(by_sorted.get(ftok_sorted)),
                       "top_candidates": scored[:5]})
    win = match_index(idx, q)
    return jsonify({
        "query": q, "forms": detail,
        "winner": None if not win else {
            "index": win["index"], "row": idx["rows"][win["index"]]["row"],
            "name": idx["rows"][win["index"]]["name"],
            "score": win["score"], "matchType": win["matchType"]},
        "index_size": len(idx["rows"]),
    }), 200


@app.get("/health")
def health():
    return jsonify({"ok": True})


@app.get("/healthz")
def healthz():
    try:
        idx = get_index()
        return jsonify({"ok": True, "rows": len(idx["rows"]), "path": XLSX_PATH,
                        "sheet": idx.get("sheet", SHEET_NAME)}), 200
    except FileNotFoundError as e:
        return jsonify({"ok": False, "error": str(e), "path": XLSX_PATH}), 503
    except Exception as e:
        return jsonify({"ok": False,
                        "error": f"{type(e).__name__}: {e}",
                        "path": XLSX_PATH}), 503


@app.get("/refresh")
@app.post("/refresh")
def refresh():
    global XLSX_PATH, _INDEX
    if not GITHUB_RAW:
        return jsonify({"error": "XLSX_URL env var not set"}), 400
    tmp = fetch_from_github()
    if not tmp:
        return jsonify({"error": "fetch failed"}), 502
    with _load_lock:
        XLSX_PATH = tmp
        _INDEX = {"key": None, "rows": [], "by_id": {}, "by_token": {},
                  "by_prefix": {}, "by_sorted": {}, "sheet": SHEET_NAME}
    try:
        idx = get_index()
    except Exception as e:
        return _err("/refresh", e, time.time())
    return jsonify({"ok": True, "rows": len(idx["rows"]), "path": XLSX_PATH}), 200


if GITHUB_RAW:
    _tmp = fetch_from_github()
    if _tmp:
        XLSX_PATH = _tmp
        print(f"[startup] fetched workbook to {XLSX_PATH}", file=sys.stderr, flush=True)

if os.path.exists(XLSX_PATH):
    _st = os.stat(XLSX_PATH)
    print(f"[startup] workbook {XLSX_PATH} ({_st.st_size} bytes)", file=sys.stderr, flush=True)
    try:
        _idx = get_index()
        print(f"[startup] indexed {len(_idx['rows'])} rows from sheet '{_idx.get('sheet','?')}'",
              file=sys.stderr, flush=True)
    except Exception as e:
        print(f"[startup] INDEX FAILED {type(e).__name__}: {e}", file=sys.stderr, flush=True)
        traceback.print_exc(file=sys.stderr)
        sys.stderr.flush()
else:
    print(f"[startup] WORKBOOK NOT FOUND at {XLSX_PATH}", file=sys.stderr, flush=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)

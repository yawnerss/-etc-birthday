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

# Columns surfaced in every response. Password is deliberately excluded.
VISIBLE = ["Student ID", "Name", "Birthday", "Campus", "Program"]


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
        campus_i = col("Campus")
        prog_i = col("Program", "Course")

        if name_i is None:
            raise ValueError(f"header 'Name' not found. Got headers: {list(headers.keys())}")

        rows, by_id, by_token, by_prefix = [], {}, {}, {}
        for r_idx, row in enumerate(it, start=2):
            if row is None:
                continue

            def cell(i):
                if i is None or i >= len(row):
                    return None
                return row[i]

            sid_s = _cell_to_str(cell(id_i))
            nm_s = _cell_to_str(cell(name_i))
            if not sid_s and not nm_s:
                continue
            bd_s = _cell_to_str(cell(bd_i))
            bd_parsed = parse_birthday(bd_s)
            name_tokens = _tokens(nm_s) if nm_s else []
            id_folded = re.sub(r"\s+", " ", _fold(sid_s)).strip() if sid_s else ""

            i = len(rows)
            rows.append({
                "row": r_idx,
                "id": sid_s,
                "name": nm_s,
                "birthday": bd_s,
                "campus": _cell_to_str(cell(campus_i)),
                "program": _cell_to_str(cell(prog_i)),
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


def row_to_public(row):
    """Map an index row to the public field set (VISIBLE order)."""
    return {
        "Student ID": row["id"],
        "Name": row["name"],
        "Birthday": row["birthday"],
        "Campus": row["campus"],
        "Program": row["program"],
    }


def full_student(idx, i):
    row = idx["rows"][i]
    out = row_to_public(row)
    out["Month"] = _MONTH_NAMES[row["birthday_month"]] if row["birthday_month"] else ""
    out["Day"] = row["birthday_day"] if row["birthday_day"] is not None else ""
    out["_row"] = row["row"]
    return out


def student_from_index(idx, i):
    return row_to_public(idx["rows"][i])


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
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="theme-color" content="#fbfbfa">
<meta name="color-scheme" content="light dark">
<title>Lookup</title>
<style>
:root{
  --bg:#fbfbfa;--surface:#ffffff;--line:#ebebe8;--fg:#161615;--muted:#8a8a84;
  --accent:#161615;--accent-fg:#ffffff;--soft:#f3f3f0;--ok:#2f7d4f;--bad:#c0392b;
  --r:16px;
}
html[data-theme="dark"]{
  --bg:#0f0f10;--surface:#171718;--line:#262628;--fg:#ececea;--muted:#7d7d82;
  --accent:#ececea;--accent-fg:#0f0f10;--soft:#1e1e20;--ok:#5fc08a;--bad:#f0776b;
}
*{box-sizing:border-box;-webkit-tap-highlight-color:transparent;}
html,body{margin:0;background:var(--bg);color:var(--fg);}
body{
  font:15px/1.5 system-ui,-apple-system,"Segoe UI",Roboto,"Helvetica Neue",sans-serif;
  -webkit-font-smoothing:antialiased;
  padding-bottom:calc(84px + env(safe-area-inset-bottom));
}
.wrap{max-width:640px;margin:0 auto;padding:0 18px;}

/* header */
header{
  position:sticky;top:0;z-index:20;background:color-mix(in srgb,var(--bg) 88%,transparent);
  backdrop-filter:blur(12px);-webkit-backdrop-filter:blur(12px);
  padding-top:env(safe-area-inset-top);
}
.bar{display:flex;align-items:center;justify-content:space-between;height:60px;}
.brand{display:flex;align-items:center;gap:10px;font-size:17px;font-weight:600;letter-spacing:-.01em;}
.dot{width:8px;height:8px;border-radius:50%;background:var(--muted);transition:background .2s;}
.dot.ok{background:var(--ok);}
.dot.bad{background:var(--bad);}
.icon-btn{
  width:44px;height:44px;border-radius:50%;border:0;background:transparent;color:var(--fg);
  font-size:18px;display:grid;place-items:center;cursor:pointer;
}
.icon-btn:active{background:var(--soft);}

/* panels */
.panel{display:none;padding-top:8px;animation:fade .18s ease;}
.panel.on{display:block;}
@keyframes fade{from{opacity:0;transform:translateY(4px);}to{opacity:1;transform:none;}}
.lead{font-size:26px;font-weight:650;letter-spacing:-.02em;margin:8px 0 18px;line-height:1.2;}
.lead span{color:var(--muted);font-weight:500;}

/* inputs: 16px to stop Android/Chrome zoom, 52px tall for thumbs */
input,select,textarea{
  width:100%;font:inherit;font-size:16px;color:var(--fg);background:var(--surface);
  border:1px solid var(--line);border-radius:var(--r);padding:0 16px;height:52px;
  outline:none;transition:border-color .15s,box-shadow .15s;appearance:none;-webkit-appearance:none;
}
input:focus,select:focus,textarea:focus{border-color:var(--fg);box-shadow:0 0 0 3px color-mix(in srgb,var(--fg) 8%,transparent);}
textarea{height:auto;min-height:170px;padding:14px 16px;resize:vertical;line-height:1.55;
  font-family:ui-monospace,SFMono-Regular,Menlo,monospace;font-size:14px;}
select{
  background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' fill='none' stroke='%238a8a84' stroke-width='2' stroke-linecap='round'%3E%3Cpath d='M1 1.5l5 5 5-5'/%3E%3C/svg%3E");
  background-repeat:no-repeat;background-position:right 18px center;padding-right:44px;
}
.stack{display:flex;flex-direction:column;gap:10px;}
.btn{
  height:52px;border-radius:var(--r);border:1px solid var(--line);background:var(--surface);color:var(--fg);
  font:inherit;font-size:15px;font-weight:600;cursor:pointer;padding:0 20px;width:100%;
  transition:transform .08s,opacity .15s;
}
.btn:active{transform:scale(.985);}
.btn.primary{background:var(--accent);color:var(--accent-fg);border-color:var(--accent);}
.btn.ghost{background:transparent;}
.btn:disabled{opacity:.45;pointer-events:none;}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:10px;}
.filebtn{display:flex;align-items:center;justify-content:center;gap:8px;height:52px;border-radius:var(--r);
  border:1px dashed var(--line);color:var(--muted);font-size:14px;cursor:pointer;background:transparent;}
.filebtn input{display:none;}

/* results */
#status{min-height:20px;font-size:13px;color:var(--muted);margin:22px 0 4px;word-break:break-word;}
#status.ok{color:var(--muted);}
#status.bad{color:var(--bad);}
.spin{display:inline-block;width:11px;height:11px;border:2px solid var(--muted);border-top-color:transparent;
  border-radius:50%;animation:sp .7s linear infinite;vertical-align:-1px;margin-right:8px;}
@keyframes sp{to{transform:rotate(360deg);}}
.sum{display:flex;gap:6px 14px;flex-wrap:wrap;font-size:13px;color:var(--muted);margin:10px 0 12px;}
.sum b{color:var(--fg);font-weight:600;}
.h3{font-size:12px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:var(--muted);margin:26px 0 10px;}
.list{background:var(--surface);border:1px solid var(--line);border-radius:var(--r);overflow:hidden;}
.item{padding:14px 16px;border-bottom:1px solid var(--line);display:flex;gap:12px;align-items:flex-start;}
.item:last-child{border-bottom:0;}
.item .body{flex:1;min-width:0;}
.item .nm{font-weight:600;font-size:15px;overflow-wrap:anywhere;}
.item .mt{font-size:13px;color:var(--muted);margin-top:2px;overflow-wrap:anywhere;}
.badge{flex:none;min-width:46px;text-align:center;padding:5px 9px;border-radius:12px;background:var(--soft);
  font-size:12px;font-weight:600;line-height:1.25;}
.badge small{display:block;font-size:10px;font-weight:500;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;}
.badge.hot{background:var(--accent);color:var(--accent-fg);}
.badge.hot small{color:inherit;opacity:.7;}
.tag{display:inline-block;font-size:11px;color:var(--muted);background:var(--soft);padding:1px 8px;border-radius:999px;margin-top:6px;}
.empty{color:var(--muted);text-align:center;padding:36px 12px;font-size:14px;}
.err{border:1px solid var(--bad);border-radius:var(--r);padding:12px 14px;color:var(--bad);font-size:13px;margin-bottom:12px;word-break:break-word;}
.kv{background:var(--surface);border:1px solid var(--line);border-radius:var(--r);padding:6px 16px;}
.kv div{display:flex;justify-content:space-between;gap:16px;padding:11px 0;border-bottom:1px solid var(--line);font-size:14px;}
.kv div:last-child{border-bottom:0;}
.kv span:first-child{color:var(--muted);flex:none;}
.kv span:last-child{text-align:right;font-weight:500;overflow-wrap:anywhere;}
.more{margin-top:12px;}

/* bottom nav (thumb zone) */
nav{
  position:fixed;left:0;right:0;bottom:0;z-index:30;background:color-mix(in srgb,var(--bg) 90%,transparent);
  backdrop-filter:blur(14px);-webkit-backdrop-filter:blur(14px);border-top:1px solid var(--line);
  padding-bottom:env(safe-area-inset-bottom);
}
nav .in{display:grid;grid-template-columns:repeat(3,1fr);max-width:640px;margin:0 auto;}
nav button{
  height:64px;border:0;background:transparent;color:var(--muted);font:inherit;font-size:11px;font-weight:600;
  letter-spacing:.02em;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:4px;cursor:pointer;
}
nav button .i{font-size:20px;line-height:1;}
nav button.on{color:var(--fg);}
nav button:active{background:var(--soft);}
@media (prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important;}}
</style>
</head>
<body>
<header>
  <div class="wrap bar">
    <div class="brand"><span class="dot" id="dot"></span><span>Lookup</span><span id="count" style="font-size:12px;color:var(--muted);font-weight:500"></span></div>
    <button class="icon-btn" id="theme" aria-label="Toggle theme">◐</button>
  </div>
</header>

<main class="wrap">

  <section class="panel on" id="p-find">
    <h1 class="lead">Find a student<br><span>by name or ID</span></h1>
    <div class="stack">
      <input id="single" placeholder="Name or student ID" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false" enterkeyhint="search">
      <button class="btn primary" id="singleBtn">Search</button>
    </div>
  </section>

  <section class="panel" id="p-month">
    <h1 class="lead">Birthdays<br><span>by month</span></h1>
    <div class="stack">
      <select id="month" aria-label="Month"></select>
      <button class="btn primary" id="scanBtn">Show this month</button>
      <button class="btn ghost" id="allBtn">Browse everyone</button>
    </div>
  </section>

  <section class="panel" id="p-list">
    <h1 class="lead">Class list<br><span>check birthdays</span></h1>
    <div class="stack">
      <textarea id="paste" wrap="off" placeholder="One per line&#10;BALINO&#9;HERIE-JAY&#10;IGARTA ALVIN&#10;26-05959" autocapitalize="off" autocorrect="off" spellcheck="false"></textarea>
      <label class="filebtn"><input type="file" id="file" accept=".txt,text/plain"><span id="fileLbl">＋ Attach .txt file</span></label>
      <div class="row2">
        <button class="btn ghost" id="clearBtn">Clear</button>
        <button class="btn primary" id="checkBtn">Check</button>
      </div>
    </div>
  </section>

  <div id="status"></div>
  <div id="out"></div>
</main>

<nav><div class="in">
  <button data-tab="find" class="on"><span class="i">⌕</span>Find</button>
  <button data-tab="month"><span class="i">◷</span>Month</button>
  <button data-tab="list"><span class="i">☰</span>List</button>
</div></nav>

<script>
const $=id=>document.getElementById(id);
const STATUS=$('status'),OUT=$('out'),PASTE=$('paste'),SINGLE=$('single');
const MONTHS=["","January","February","March","April","May","June","July","August","September","October","November","December"];
const BATCH_SIZE=400,RENDER_CAP=200,LS_PASTE='bday-paste';
let LAST_ALL=null;

function el(t,c,x){const e=document.createElement(t);if(c)e.className=c;if(x!==undefined&&x!==null)e.textContent=x;return e;}
function safe(f){try{return f();}catch(_){return null;}}

/* theme */
function applyTheme(t){
  document.documentElement.setAttribute('data-theme',t);
  $('theme').textContent=t==='dark'?'☀':'◐';
  const m=document.querySelector('meta[name=theme-color]');
  if(m)m.setAttribute('content',t==='dark'?'#0f0f10':'#fbfbfa');
  safe(()=>localStorage.setItem('bday-theme',t));
}
applyTheme(safe(()=>localStorage.getItem('bday-theme'))||(matchMedia('(prefers-color-scheme: dark)').matches?'dark':'light'));
$('theme').onclick=()=>applyTheme(document.documentElement.getAttribute('data-theme')==='dark'?'light':'dark');

/* tabs */
document.querySelectorAll('nav button').forEach(b=>b.onclick=()=>{
  document.querySelectorAll('nav button').forEach(x=>x.classList.toggle('on',x===b));
  document.querySelectorAll('.panel').forEach(p=>p.classList.toggle('on',p.id==='p-'+b.dataset.tab));
  window.scrollTo({top:0,behavior:'smooth'});
});

/* month select */
(function(){
  const cur=new Date().getMonth()+1,sel=$('month');
  for(let i=1;i<=12;i++){
    const o=document.createElement('option');o.value=i;
    o.textContent=MONTHS[i]+(i===cur?'  ·  now':'');sel.appendChild(o);
  }
  sel.value=cur;
})();
const monthSel=()=>parseInt($('month').value,10);

/* paste persistence */
const saved=safe(()=>localStorage.getItem(LS_PASTE));if(saved)PASTE.value=saved;
let _t=null;
PASTE.addEventListener('input',()=>{clearTimeout(_t);_t=setTimeout(()=>safe(()=>localStorage.setItem(LS_PASTE,PASTE.value)),400);});
$('file').onchange=e=>{$('fileLbl').textContent=e.target.files[0]?('✓ '+e.target.files[0].name):'＋ Attach .txt file';};

/* server status */
async function checkServer(){
  const d=$('dot');d.className='dot';
  try{
    const j=await (await fetch('/healthz',{cache:'no-store'})).json();
    d.className='dot '+(j.ok?'ok':'bad');
    $('count').textContent=j.ok?(j.rows+' students'):'';
    d.title=j.ok?'':(j.error||'error');
  }catch(e){d.className='dot bad';$('count').textContent='offline';}
}
checkServer();

function setStatus(m,k){
  STATUS.innerHTML='';
  if(k==='load')STATUS.appendChild(el('span','spin'));
  STATUS.appendChild(document.createTextNode(m||''));
  STATUS.className=k==='load'?'':(k||'');
}
function setBusy(b){['checkBtn','scanBtn','allBtn','singleBtn'].forEach(id=>$(id).disabled=b);}

/* components */
function meta(s){
  const a=[s['Student ID'],s['Birthday']].filter(Boolean).join(' · ');
  const b=[s['Campus'],s['Program']].filter(Boolean).join(' · ');
  return [a,b];
}
function badge(month,day,hot){
  if(!day&&!month)return null;
  const b=el('div','badge'+(hot?' hot':''));
  b.appendChild(document.createTextNode(day||'–'));
  b.appendChild(el('small',null,month?MONTHS[month].slice(0,3):''));
  return b;
}
function item(s,opts){
  opts=opts||{};
  const it=el('div','item');
  const body=el('div','body');
  body.appendChild(el('div','nm',s['Name']||opts.fallback||'—'));
  const m=meta(s);
  if(m[0])body.appendChild(el('div','mt',m[0]));
  if(m[1])body.appendChild(el('div','mt',m[1]));
  if(opts.tag)body.appendChild(el('span','tag',opts.tag));
  it.appendChild(body);
  const b=badge(opts.month,opts.day,opts.hot);
  if(b)it.appendChild(b);
  return it;
}
function listOf(rows,fn,cap){
  const box=el('div','list');
  if(!rows.length){box.appendChild(el('div','empty','Nothing here.'));return box;}
  (cap&&rows.length>cap?rows.slice(0,cap):rows).forEach(r=>box.appendChild(fn(r)));
  if(cap&&rows.length>cap)box.appendChild(el('div','empty','Showing '+cap+' of '+rows.length+' — use CSV for all.'));
  return box;
}
function sum(parts){
  const s=el('div','sum');
  parts.forEach(p=>{const sp=el('span');sp.innerHTML=p;s.appendChild(sp);});
  return s;
}
function esc(x){return String(x==null?'':x).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}
function csvDownload(rows,name){
  const csv=rows.map(r=>r.map(v=>'"'+String(v==null?'':v).replace(/"/g,'""')+'"').join(',')).join('\n');
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'}));
  a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);
}
function csvBtn(label,fn){const b=el('button','btn ghost more',label);b.onclick=fn;return b;}

/* renderers */
function renderSingle(d){
  OUT.innerHTML='';
  const s=d.student;
  OUT.appendChild(sum(['match <b>'+esc(d.matchType)+'</b>','score <b>'+esc(d.score)+'</b>','row <b>'+esc(s._row)+'</b>']));
  const kv=el('div','kv');
  [['Student ID',s['Student ID']],['Name',s['Name']],['Birthday',s['Birthday']],['Campus',s['Campus']],
   ['Program',s['Program']],['Month',s['Month']],['Day',s['Day']]].forEach(([k,v])=>{
    const r=el('div');r.appendChild(el('span',null,k));r.appendChild(el('span',null,v===''||v==null?'—':v));kv.appendChild(r);
  });
  OUT.appendChild(kv);
  OUT.appendChild(csvBtn('Copy details',()=>navigator.clipboard&&navigator.clipboard.writeText(JSON.stringify(s,null,2))));
}
function renderMonth(d){
  OUT.innerHTML='';
  OUT.appendChild(sum(['<b>'+d.count+'</b> in <b>'+esc(d.monthName)+'</b>']));
  const mi=MONTHS.indexOf(d.monthName);
  OUT.appendChild(listOf(d.rows,r=>item(r,{month:mi,day:r.day,hot:true}),RENDER_CAP));
}
function renderAll(d){
  LAST_ALL=d;OUT.innerHTML='';
  OUT.appendChild(sum(['<b>'+d.count+'</b> students']));
  const f=el('input');f.placeholder='Filter name, ID, campus, month…';f.setAttribute('autocapitalize','off');
  OUT.appendChild(f);
  const holder=el('div');holder.style.marginTop='12px';OUT.appendChild(holder);
  const cols=['Student ID','Name','Birthday','Campus','Program','Month','Day'];
  const draw=q=>{
    q=(q||'').trim().toLowerCase();
    const rows=q?d.rows.filter(r=>cols.some(c=>String(r[c]).toLowerCase().includes(q))):d.rows;
    holder.innerHTML='';
    holder.appendChild(sum(['showing <b>'+Math.min(rows.length,RENDER_CAP)+'</b> of <b>'+rows.length+'</b>']));
    holder.appendChild(listOf(rows,r=>item(r,{month:MONTHS.indexOf(r.Month),day:r.Day}),RENDER_CAP));
  };
  f.oninput=()=>draw(f.value);draw('');
  OUT.appendChild(csvBtn('Download CSV',()=>{
    const rows=[cols];d.rows.forEach(r=>rows.push(cols.map(c=>r[c])));csvDownload(rows,'students-all.csv');
  }));
}
function renderList(a){
  OUT.innerHTML='';
  const p=['<b>'+a.total+'</b> lines','<b>'+a.matchedCount+'</b> matched','<b>'+a.thisMonthCount+'</b> in <b>'+esc(a.monthName||'—')+'</b>'];
  if(a.duplicates&&a.duplicates.length)p.push('<b>'+a.duplicates.length+'</b> duplicates skipped');
  if(a.unmatchedCount)p.push('<b>'+a.unmatchedCount+'</b> unmatched');
  OUT.appendChild(sum(p));
  if(a.lastError)OUT.appendChild(el('div','err','Server error: '+a.lastError));
  if(!a.matchedCount&&!a.unmatchedCount)return;

  const mk=(e,hot)=>item(e.student||{},{fallback:e.line,month:e.birthdayMonth,day:e.birthdayDay,hot:hot,
    tag:e.matchType==='inline'?'date from line':(e.matched&&e.matchType!=='id'&&e.matchType!=='exact'?e.matchType+' · '+e.score:null)});

  OUT.appendChild(el('div','h3','Birthday in '+(a.monthName||'—')+' · '+a.thisMonth.length));
  OUT.appendChild(listOf(a.thisMonth,e=>mk(e,true),RENDER_CAP));
  if(a.matchedOther.length){
    OUT.appendChild(el('div','h3','Other months · '+a.matchedOther.length));
    OUT.appendChild(listOf(a.matchedOther,e=>mk(e,false),RENDER_CAP));
  }
  if(a.unmatched.length){
    OUT.appendChild(el('div','h3','Unmatched · '+a.unmatched.length));
    OUT.appendChild(listOf(a.unmatched,u=>{
      const it=el('div','item'),b=el('div','body');
      b.appendChild(el('div','nm',u.line));b.appendChild(el('div','mt',u.error||'no match'));
      it.appendChild(b);return it;
    },RENDER_CAP));
  }
  OUT.appendChild(csvBtn('Download CSV',()=>{
    const rows=[['Group','Student ID','Name','Birthday','Campus','Program','Match','Score']];
    const push=(g,e)=>{const s=e.student||{};rows.push([g,s['Student ID']||'',s['Name']||e.line||'',s['Birthday']||'',s['Campus']||'',s['Program']||'',e.matchType||'',e.score||'']);};
    a.thisMonth.forEach(e=>push('thisMonth',e));a.matchedOther.forEach(e=>push('other',e));
    a.unmatched.forEach(e=>rows.push(['unmatched','',e.line||'','','','','','']));
    csvDownload(rows,'birthdays-'+(a.monthName||'month')+'.csv');
  }));
}

/* network */
async function withTimeout(url,opts,ms){
  const c=new AbortController(),t=setTimeout(()=>c.abort(),ms);
  try{return await fetch(url,{...opts,signal:c.signal});}finally{clearTimeout(t);}
}
async function readErr(r){try{return (await r.json()).error||('HTTP '+r.status);}catch(_){return 'HTTP '+r.status;}}
const fail=e=>setStatus(e.name==='AbortError'?'Timed out.':'Request failed: '+e.message,'bad');

async function lookupSingle(){
  const q=SINGLE.value.trim();
  if(!q){setStatus('Enter a name or student ID.','bad');return;}
  SINGLE.blur();setBusy(true);setStatus('Searching…','load');
  try{
    const r=await withTimeout('/student?q='+encodeURIComponent(q),{},15000);
    if(r.status===404){OUT.innerHTML='';setStatus('No match for “'+q+'”.','bad');return;}
    if(!r.ok){setStatus('Error: '+(await readErr(r)),'bad');return;}
    const j=await r.json();
    setStatus(j.elapsedMs+' ms','ok');renderSingle(j);
  }catch(e){fail(e);}finally{setBusy(false);checkServer();}
}
async function scanWorkbook(){
  setBusy(true);setStatus('Loading…','load');
  try{
    const r=await withTimeout('/birthdays?month='+monthSel(),{},15000);
    if(!r.ok){setStatus('Error: '+(await readErr(r)),'bad');return;}
    const j=await r.json();setStatus(j.elapsedMs+' ms','ok');renderMonth(j);
  }catch(e){fail(e);}finally{setBusy(false);checkServer();}
}
async function scanAll(){
  setBusy(true);setStatus('Loading everyone…','load');
  try{
    const r=await withTimeout('/all',{},30000);
    if(!r.ok){setStatus('Error: '+(await readErr(r)),'bad');return;}
    const j=await r.json();setStatus(j.elapsedMs+' ms','ok');renderAll(j);
  }catch(e){fail(e);}finally{setBusy(false);checkServer();}
}
async function checkList(){
  const raw=PASTE.value,file=$('file').files[0];
  if(!raw.trim()&&!file){setStatus('Paste names or attach a txt file.','bad');return;}
  const month=monthSel(),t0=performance.now();
  setBusy(true);

  if(file){
    setStatus('Uploading…','load');
    try{
      const fd=new FormData();fd.append('file',file);fd.append('month',String(month));
      const r=await withTimeout('/birthdays/list',{method:'POST',body:fd},120000);
      if(!r.ok){setStatus('Error: '+(await readErr(r)),'bad');return;}
      const j=await r.json();
      setStatus(Math.round(performance.now()-t0)+' ms','ok');renderList(j);
    }catch(e){fail(e);}finally{setBusy(false);checkServer();}
    return;
  }

  const lines=raw.split('\n').map(s=>s.trim()).filter(Boolean);
  const batches=[];for(let i=0;i<lines.length;i+=BATCH_SIZE)batches.push(lines.slice(i,i+BATCH_SIZE));
  const agg={monthName:null,total:0,duplicates:[],thisMonth:[],matchedOther:[],unmatched:[],failed:0,lastError:null};
  try{
    for(let b=0;b<batches.length;b++){
      setStatus('Batch '+(b+1)+' of '+batches.length+'…','load');
      try{
        const r=await withTimeout('/birthdays/list',{method:'POST',headers:{'Content-Type':'application/json'},
          body:JSON.stringify({text:batches[b].join('\n'),month})},60000);
        if(!r.ok){agg.failed++;if(!agg.lastError)agg.lastError=await readErr(r);continue;}
        let j;try{j=await r.json();}catch(_){agg.failed++;if(!agg.lastError)agg.lastError='bad JSON';continue;}
        agg.monthName=j.monthName;agg.total+=j.total||0;
        agg.duplicates.push(...(j.duplicates||[]));agg.thisMonth.push(...(j.thisMonth||[]));
        agg.matchedOther.push(...(j.matchedOther||[]));agg.unmatched.push(...(j.unmatched||[]));
      }catch(e){agg.failed++;if(!agg.lastError)agg.lastError=String(e.message||e);}
    }
    agg.thisMonth.sort((a,b)=>((a.birthdayDay||99)-(b.birthdayDay||99))||String((a.student||{}).Name||'').localeCompare(String((b.student||{}).Name||'')));
    agg.matchedOther.sort((a,b)=>((a.birthdayMonth||99)-(b.birthdayMonth||99))||((a.birthdayDay||99)-(b.birthdayDay||99)));
    agg.matchedCount=agg.thisMonth.length+agg.matchedOther.length;
    agg.thisMonthCount=agg.thisMonth.length;agg.unmatchedCount=agg.unmatched.length;
    setStatus(Math.round(performance.now()-t0)+' ms'+(agg.failed?' · '+agg.failed+' batch(es) failed':''),agg.failed?'bad':'ok');
    renderList(agg);
  }finally{setBusy(false);checkServer();}
}
function clearAll(){
  PASTE.value='';safe(()=>localStorage.removeItem(LS_PASTE));
  $('file').value='';$('fileLbl').textContent='＋ Attach .txt file';
  SINGLE.value='';OUT.innerHTML='';setStatus('');
}

$('singleBtn').onclick=lookupSingle;
$('scanBtn').onclick=scanWorkbook;
$('allBtn').onclick=scanAll;
$('checkBtn').onclick=checkList;
$('clearBtn').onclick=clearAll;
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
        "visible_columns": VISIBLE,
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
        if idx["rows"]:
            out["index_sample"] = row_to_public(idx["rows"][0])
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
            item = row_to_public(row)
            item["_row"] = row["row"]
            item["day"] = row["birthday_day"]
            out.append(item)
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
            item = row_to_public(row)
            item["Month"] = _MONTH_NAMES[row["birthday_month"]] if row["birthday_month"] else ""
            item["Day"] = row["birthday_day"] if row["birthday_day"] is not None else ""
            item["_row"] = row["row"]
            out.append(item)

        q = (request.args.get("q") or "").strip().lower()
        if q:
            fields = VISIBLE + ["Month"]
            out = [r for r in out if any(q in str(r.get(f, "")).lower() for f in fields)]

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
                                "Birthday": inline_bd.strftime("%B %d, %Y"),
                                "Campus": "", "Program": ""},
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
            "student": row_to_public(idx["rows"][win["index"]]),
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
                        "sheet": idx.get("sheet", SHEET_NAME),
                        "columns": VISIBLE}), 200
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
        print(f"[startup] indexed {len(_idx['rows'])} rows from sheet '{_idx.get('sheet','?')}', "
              f"columns visible={VISIBLE}", file=sys.stderr, flush=True)
    except Exception as e:
        print(f"[startup] INDEX FAILED {type(e).__name__}: {e}", file=sys.stderr, flush=True)
        traceback.print_exc(file=sys.stderr)
        sys.stderr.flush()
else:
    print(f"[startup] WORKBOOK NOT FOUND at {XLSX_PATH}", file=sys.stderr, flush=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)

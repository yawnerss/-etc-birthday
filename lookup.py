#!/usr/bin/env python3
import os
import re
import sys
import threading
import time
import traceback
import unicodedata
import urllib.request
from collections import Counter
from datetime import datetime, date, timedelta, timezone
from difflib import SequenceMatcher

from flask import Flask, jsonify, request
from openpyxl import load_workbook

GITHUB_RAW = os.environ.get("XLSX_URL", "").strip()

# Server runs in UTC on most hosts; this keeps "today" / age correct for your timezone.
# Philippines / Singapore = 8. Override with env TZ_OFFSET_HOURS.
try:
    TZ_OFFSET_HOURS = float(os.environ.get("TZ_OFFSET_HOURS", "8"))
except ValueError:
    TZ_OFFSET_HOURS = 8.0

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
CANDIDATE_LIMIT = 80

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


# ----------------------------------------------------------------------------
# time / age helpers
# ----------------------------------------------------------------------------
def _today():
    return (datetime.now(timezone.utc) + timedelta(hours=TZ_OFFSET_HOURS)).date()


def _age_on(y, m, d, today):
    return today.year - y - ((today.month, today.day) < (m, d))


def _next_birthday(m, d, today):
    for yr in (today.year, today.year + 1):
        try:
            nb = date(yr, m, d)
        except ValueError:  # Feb 29 on a non-leap year
            nb = date(yr, 2, 28)
        if nb >= today:
            return nb
    return None


_ZODIAC = [
    (1, 20, "Capricorn"), (2, 19, "Aquarius"), (3, 21, "Pisces"), (4, 20, "Aries"),
    (5, 21, "Taurus"), (6, 21, "Gemini"), (7, 23, "Cancer"), (8, 23, "Leo"),
    (9, 23, "Virgo"), (10, 23, "Libra"), (11, 22, "Scorpio"), (12, 22, "Sagittarius"),
]


def _zodiac(m, d):
    for em, ed, name in _ZODIAC:
        if (m, d) < (em, ed):
            return name
    return "Capricorn"


def _derived(m, d, y, today=None):
    """Age / days-to-next-birthday / zodiac from month, day and (optional) year."""
    today = today or _today()
    out = {"Age": "", "DaysToBday": "", "Zodiac": ""}
    if not m or not d:
        return out
    try:
        nb = _next_birthday(m, d, today)
        if nb:
            out["DaysToBday"] = (nb - today).days
        out["Zodiac"] = _zodiac(m, d)
        if y:
            age = _age_on(y, m, d, today)
            if 0 <= age <= 120:
                out["Age"] = age
    except Exception:
        pass
    return out


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
                # 1900 is the placeholder year used when a birthday has no year
                "birthday_year": (bd_parsed.year if bd_parsed and bd_parsed.year != 1900 else None),
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


# ----------------------------------------------------------------------------
# candidate search: "james" -> every student with a James in their name
# ----------------------------------------------------------------------------
def _tok_match(qt, ct):
    if qt == ct:
        return 1.0
    if len(qt) >= 2 and ct.startswith(qt):
        return 0.92
    if len(qt) >= 3 and qt in ct:
        return 0.80
    if len(qt) >= 3 and abs(len(qt) - len(ct)) <= 2:
        sm = SequenceMatcher(None, qt, ct)
        if sm.real_quick_ratio() >= 0.8 and sm.quick_ratio() >= 0.8:
            r = sm.ratio()
            if r >= 0.8:
                return r * 0.85
    return 0.0


def find_candidates(idx, q, limit=CANDIDATE_LIMIT):
    """Every row where each query token matches some name token.
    Returns (list of (score, index) best-first, total_found)."""
    qt = _tokens(q)
    if not qt:
        return [], 0
    rows = idx["rows"]
    found = []
    for i, r in enumerate(rows):
        ct = r["name_tokens"]
        if not ct:
            continue
        total = 0.0
        ok = True
        for t in qt:
            best = 0.0
            for c in ct:
                m = _tok_match(t, c)
                if m > best:
                    best = m
                    if best == 1.0:
                        break
            if best == 0.0:
                ok = False
                break
            total += best
        if ok:
            found.append((total / len(qt), i))
    found.sort(key=lambda x: (-x[0], rows[x[1]]["name"].lower()))
    return found[:limit], len(found)


def row_to_public(row, today=None):
    """Map an index row to the public field set (VISIBLE order + derived fields)."""
    d = _derived(row["birthday_month"], row["birthday_day"], row["birthday_year"], today)
    return {
        "Student ID": row["id"],
        "Name": row["name"],
        "Birthday": row["birthday"],
        "Campus": row["campus"],
        "Program": row["program"],
        "Age": d["Age"],
        "Zodiac": d["Zodiac"],
        "DaysToBday": d["DaysToBday"],
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
        m = _today().month
        return m, _MONTH_NAMES[m]
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
<html lang="en" data-theme="green">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="theme-color" content="#09090a">
<meta name="color-scheme" content="dark">
<title>lookup</title>
<style>
:root{
  --g:#3dff8f;--bg:#09090a;--fg:#dfe3df;--muted:#6c746e;--line:rgba(255,255,255,.075);
  --soft:rgba(255,255,255,.035);--bad:#ff6b6b;
}
html[data-theme="amber"]{--g:#ffb454;}
html[data-theme="red"]{--g:#ff5c6c;--bad:#ffd166;}
html[data-theme="cyan"]{--g:#4fd8ff;}
*{box-sizing:border-box;-webkit-tap-highlight-color:transparent;}
html,body{margin:0;background:var(--bg);color:var(--fg);}
body{
  font:14px/1.55 ui-monospace,"SF Mono",SFMono-Regular,Menlo,Consolas,"Liberation Mono",monospace;
  -webkit-font-smoothing:antialiased;min-height:100vh;
  padding-bottom:calc(72px + env(safe-area-inset-bottom));
  background:
    radial-gradient(900px 420px at 50% -120px,color-mix(in srgb,var(--g) 9%,transparent),transparent 70%),
    var(--bg);
  background-attachment:fixed;
}
.wrap{max-width:620px;margin:0 auto;padding:0 22px;}
::selection{background:var(--g);color:#000;}

/* top progress line */
#prog{position:fixed;left:0;top:0;height:2px;width:0;background:var(--g);z-index:50;opacity:0;
  box-shadow:0 0 12px var(--g);transition:width .2s ease,opacity .3s ease;}

/* header */
header{position:sticky;top:0;z-index:20;padding-top:env(safe-area-inset-top);
  background:linear-gradient(var(--bg) 60%,transparent);}
.bar{display:flex;align-items:center;justify-content:space-between;height:60px;}
.brand{display:flex;align-items:center;gap:10px;font-size:13px;color:var(--muted);letter-spacing:.04em;}
.brand b{color:var(--fg);font-weight:600;}
.brand b::after{content:"_";color:var(--g);animation:blink 1.1s steps(1) infinite;}
@keyframes blink{50%{opacity:0;}}
.dot{width:6px;height:6px;border-radius:50%;background:var(--muted);transition:background .3s,box-shadow .3s;}
.dot.ok{background:var(--g);box-shadow:0 0 8px var(--g);}
.dot.bad{background:var(--bad);}
#theme{width:30px;height:30px;border:0;background:transparent;display:grid;place-items:center;cursor:pointer;border-radius:50%;}
#theme i{width:10px;height:10px;border-radius:50%;background:var(--g);box-shadow:0 0 10px var(--g);display:block;transition:transform .2s;}
#theme:active i{transform:scale(1.5);}

/* panels */
.panel{display:none;padding-top:18px;}
.panel.on{display:block;animation:rise .35s cubic-bezier(.2,.7,.2,1);}
@keyframes rise{from{opacity:0;transform:translateY(8px);}to{opacity:1;transform:none;}}
.lead{font-size:30px;font-weight:300;letter-spacing:-.02em;margin:0 0 4px;color:var(--fg);}
.lead::after{content:"_";color:var(--g);}
.sub{color:var(--muted);font-size:12.5px;margin:0 0 26px;}
.sub b{color:var(--g);font-weight:500;}

/* inputs */
input,select,textarea{
  width:100%;font:inherit;font-size:16px;color:var(--fg);background:transparent;border:0;
  border-bottom:1px solid var(--line);border-radius:0;padding:0;height:52px;outline:none;
  transition:border-color .25s;appearance:none;-webkit-appearance:none;caret-color:var(--g);
}
input.big{font-size:20px;}
input::placeholder,textarea::placeholder{color:var(--muted);opacity:.7;}
input:focus,select:focus{border-bottom-color:var(--g);}
textarea{height:auto;min-height:170px;padding:14px 16px;border:0;border-radius:12px;background:var(--soft);
  resize:vertical;line-height:1.6;font-size:14px;transition:box-shadow .25s;}
textarea:focus{box-shadow:0 0 0 1px var(--g);}
select{background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6' fill='none' stroke='%236c746e' stroke-width='1.6' stroke-linecap='round'%3E%3Cpath d='M1 1l4 4 4-4'/%3E%3C/svg%3E");
  background-repeat:no-repeat;background-position:right 2px center;padding-right:24px;}
select option{background:#111;color:#ddd;}
.stack{display:flex;flex-direction:column;gap:14px;}
.btn{
  height:46px;border-radius:999px;border:1px solid var(--line);background:transparent;color:var(--fg);
  font:inherit;font-size:13px;font-weight:600;letter-spacing:.04em;cursor:pointer;padding:0 22px;width:100%;
  transition:transform .12s,background .2s,box-shadow .2s,border-color .2s;
}
.btn:hover{border-color:color-mix(in srgb,var(--g) 55%,transparent);}
.btn:active{transform:scale(.985);}
.btn.primary{background:var(--g);color:#04110a;border-color:var(--g);}
.btn.primary:hover{box-shadow:0 0 22px color-mix(in srgb,var(--g) 40%,transparent);}
.btn.ghost{color:var(--muted);}
.btn.sm{height:38px;font-size:12px;width:auto;margin-top:18px;}
.btn:disabled{opacity:.35;pointer-events:none;}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:12px;}
.filebtn{display:flex;align-items:center;justify-content:center;height:44px;border-radius:999px;
  border:1px dashed var(--line);color:var(--muted);font-size:12.5px;cursor:pointer;transition:border-color .2s,color .2s;}
.filebtn:hover{border-color:var(--g);color:var(--g);}
.filebtn input{display:none;}
.hint{color:var(--muted);font-size:12px;margin-top:16px;text-align:center;}

/* status */
#status{min-height:22px;font-size:12px;color:var(--muted);margin:22px 0 6px;word-break:break-word;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
#status.ok{color:var(--g);}
#status.bad{color:var(--bad);white-space:normal;}
#out{animation:rise .35s cubic-bezier(.2,.7,.2,1);}

/* results */
.sum{display:flex;gap:4px 16px;flex-wrap:wrap;font-size:12px;color:var(--muted);margin:6px 0 10px;}
.sum b{color:var(--fg);font-weight:600;}
.h3{font-size:11px;letter-spacing:.16em;text-transform:uppercase;color:var(--muted);margin:34px 0 6px;}
.h3 em{font-style:normal;color:var(--g);margin-left:8px;}
.list{border-top:1px solid var(--line);}
.item{padding:14px 2px;border-bottom:1px solid var(--line);display:flex;gap:16px;align-items:center;justify-content:space-between;
  transition:background .2s,padding .2s;}
.item.click{cursor:pointer;}
.item.click:hover{background:var(--soft);padding-left:10px;padding-right:10px;}
.item .body{min-width:0;flex:1;}
.item .nm{font-weight:500;font-size:14.5px;overflow-wrap:anywhere;}
.item .mt{font-size:12px;color:var(--muted);overflow-wrap:anywhere;}
.item .side{flex:none;text-align:right;line-height:1.2;}
.item .age{font-size:20px;font-weight:300;color:var(--g);font-variant-numeric:tabular-nums;}
.item .age small{font-size:10px;color:var(--muted);margin-left:3px;letter-spacing:.06em;}
.item .dt{font-size:11px;color:var(--muted);margin-top:3px;}
.pill{display:inline-block;font-size:10px;letter-spacing:.08em;color:var(--g);border:1px solid color-mix(in srgb,var(--g) 45%,transparent);
  padding:0 7px;border-radius:999px;margin-top:5px;margin-right:5px;line-height:1.6;text-transform:uppercase;}
.pill.hot{background:var(--g);color:#04110a;border-color:var(--g);font-weight:700;}
.pill.mute{color:var(--muted);border-color:var(--line);}
.empty{color:var(--muted);text-align:center;padding:34px 12px;font-size:13px;}
.err{border:1px solid var(--bad);border-radius:12px;padding:12px 14px;color:var(--bad);font-size:12px;margin-bottom:12px;word-break:break-word;}
.kv{border-top:1px solid var(--line);}
.kv div{display:flex;justify-content:space-between;gap:20px;padding:12px 2px;border-bottom:1px solid var(--line);font-size:13.5px;}
.kv span:first-child{color:var(--muted);flex:none;font-size:11px;letter-spacing:.12em;text-transform:uppercase;padding-top:2px;}
.kv span:last-child{text-align:right;overflow-wrap:anywhere;}
.kv .accent{color:var(--g);}
.dossier-name{font-size:26px;font-weight:300;letter-spacing:-.01em;margin:10px 0 18px;overflow-wrap:anywhere;}
.dossier-name small{display:block;font-size:11px;letter-spacing:.16em;color:var(--g);text-transform:uppercase;margin-bottom:6px;}

/* stats */
.cards{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin:8px 0 0;}
.card b{display:block;font-size:26px;font-weight:300;color:var(--g);font-variant-numeric:tabular-nums;}
.card small{font-size:10px;color:var(--muted);letter-spacing:.14em;text-transform:uppercase;}
.bars{border-top:1px solid var(--line);padding-top:8px;}
.brow{display:grid;grid-template-columns:70px 1fr 40px;gap:12px;align-items:center;padding:7px 0;font-size:12.5px;transition:opacity .2s;}
.bars.wide .brow{grid-template-columns:120px 1fr 40px;}
.brow[style*="pointer"]:hover .bk{color:var(--g);}
.brow .bk{color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;transition:color .2s;}
.brow .bt{height:3px;background:var(--soft);border-radius:2px;overflow:hidden;}
.brow .bt i{display:block;height:100%;background:var(--g);border-radius:2px;box-shadow:0 0 8px color-mix(in srgb,var(--g) 60%,transparent);}
.brow .bv{text-align:right;color:var(--fg);font-variant-numeric:tabular-nums;}

/* nav */
nav{position:fixed;left:0;right:0;bottom:0;z-index:30;padding-bottom:env(safe-area-inset-bottom);
  background:linear-gradient(transparent,var(--bg) 40%);}
nav .in{display:grid;grid-template-columns:repeat(4,1fr);max-width:620px;margin:0 auto;}
nav button{height:62px;border:0;background:transparent;color:var(--muted);font:inherit;font-size:12px;letter-spacing:.08em;
  cursor:pointer;position:relative;transition:color .25s;}
nav button::before{content:"";position:absolute;left:50%;top:10px;width:0;height:2px;background:var(--g);border-radius:2px;
  transform:translateX(-50%);transition:width .25s;box-shadow:0 0 8px var(--g);}
nav button.on{color:var(--g);}
nav button.on::before{width:18px;}
@media (prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important;}}
</style>
</head>
<body>
<div id="prog"></div>

<header>
  <div class="wrap bar">
    <div class="brand"><span class="dot" id="dot"></span><b>lookup</b><span id="count"></span></div>
    <button id="theme" aria-label="Change accent color"><i></i></button>
  </div>
</header>

<main class="wrap">

  <section class="panel on" id="p-find">
    <h1 class="lead">find</h1>
    <p class="sub">name, partial name or student ID. try just <b>james</b></p>
    <div class="stack">
      <input class="big" id="single" placeholder="who are you looking for" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false" enterkeyhint="search">
      <button class="btn primary" id="singleBtn">scan</button>
    </div>
  </section>

  <section class="panel" id="p-month">
    <h1 class="lead">month</h1>
    <p class="sub">every birthday in a month, with ages</p>
    <div class="stack">
      <select id="month" aria-label="Month"></select>
      <button class="btn primary" id="scanBtn">sweep month</button>
      <button class="btn ghost" id="allBtn">browse everyone</button>
    </div>
  </section>

  <section class="panel" id="p-list">
    <h1 class="lead">list</h1>
    <p class="sub">paste a class list, see who has a birthday in the selected month</p>
    <div class="stack">
      <textarea id="paste" wrap="off" placeholder="one per line&#10;BALINO&#9;HERIE-JAY&#10;IGARTA ALVIN&#10;26-05959" autocapitalize="off" autocorrect="off" spellcheck="false"></textarea>
      <label class="filebtn"><input type="file" id="file" accept=".txt,text/plain"><span id="fileLbl">attach .txt</span></label>
      <div class="row2">
        <button class="btn ghost" id="clearBtn">clear</button>
        <button class="btn primary" id="checkBtn">check</button>
      </div>
    </div>
  </section>

  <section class="panel" id="p-intel">
    <h1 class="lead">intel</h1>
    <p class="sub">today, incoming birthdays, students per campus</p>
    <div class="stack">
      <select id="days" aria-label="Look-ahead window">
        <option value="7">next 7 days</option>
        <option value="14">next 14 days</option>
        <option value="30" selected>next 30 days</option>
        <option value="60">next 60 days</option>
        <option value="90">next 90 days</option>
      </select>
      <button class="btn primary" id="statsBtn">run recon</button>
    </div>
  </section>

  <div id="status"></div>
  <div id="out"></div>
</main>

<nav><div class="in">
  <button data-tab="find" class="on">find</button>
  <button data-tab="month">month</button>
  <button data-tab="list">list</button>
  <button data-tab="intel">intel</button>
</div></nav>

<script>
const $=id=>document.getElementById(id);
const STATUS=$('status'),OUT=$('out'),PASTE=$('paste'),SINGLE=$('single');
const MONTHS=["","January","February","March","April","May","June","July","August","September","October","November","December"];
const BATCH_SIZE=400,RENDER_CAP=200,LS_PASTE='bday-paste';
const sleep=ms=>new Promise(r=>setTimeout(r,ms));

function el(t,c,x){const e=document.createElement(t);if(c)e.className=c;if(x!==undefined&&x!==null)e.textContent=x;return e;}
function safe(f){try{return f();}catch(_){return null;}}
function esc(x){return String(x==null?'':x).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}

/* theme: tiny accent dot in the header */
const THEMES=['green','amber','red','cyan'];
function applyTheme(t){
  if(THEMES.indexOf(t)<0)t='green';
  document.documentElement.setAttribute('data-theme',t);
  safe(()=>localStorage.setItem('hk-theme',t));
}
applyTheme(safe(()=>localStorage.getItem('hk-theme'))||'green');
$('theme').onclick=()=>{
  const cur=document.documentElement.getAttribute('data-theme');
  applyTheme(THEMES[(THEMES.indexOf(cur)+1)%THEMES.length]);
};

/* tabs */
let STATS_LOADED=false;
document.querySelectorAll('nav button').forEach(b=>b.onclick=()=>{
  document.querySelectorAll('nav button').forEach(x=>x.classList.toggle('on',x===b));
  document.querySelectorAll('.panel').forEach(p=>p.classList.toggle('on',p.id==='p-'+b.dataset.tab));
  window.scrollTo({top:0,behavior:'smooth'});
  if(b.dataset.tab==='intel'&&!STATS_LOADED)loadStats();
});

/* month select */
(function(){
  const cur=new Date().getMonth()+1,sel=$('month');
  for(let i=1;i<=12;i++){
    const o=document.createElement('option');o.value=i;
    o.textContent=MONTHS[i].toLowerCase()+(i===cur?'  (now)':'');sel.appendChild(o);
  }
  sel.value=cur;
})();
const monthSel=()=>parseInt($('month').value,10);

/* paste persistence */
const saved=safe(()=>localStorage.getItem(LS_PASTE));if(saved)PASTE.value=saved;
let _t=null;
PASTE.addEventListener('input',()=>{clearTimeout(_t);_t=setTimeout(()=>safe(()=>localStorage.setItem(LS_PASTE,PASTE.value)),400);});
$('file').onchange=e=>{$('fileLbl').textContent=e.target.files[0]?e.target.files[0].name:'attach .txt';};

/* server status */
async function checkServer(){
  const d=$('dot');d.className='dot';
  try{
    const j=await (await fetch('/healthz',{cache:'no-store'})).json();
    d.className='dot '+(j.ok?'ok':'bad');
    $('count').textContent=j.ok?(j.rows+' records'):'';
    d.title=j.ok?'':(j.error||'error');
  }catch(e){d.className='dot bad';$('count').textContent='offline';}
}
checkServer();

function setStatus(m,k){STATUS.textContent=m||'';STATUS.className=k||'';}
function setBusy(b){['checkBtn','scanBtn','allBtn','singleBtn','statsBtn'].forEach(id=>$(id).disabled=b);}

/* subtle scan effect: thin progress line + one rotating log line */
const SCAN_LOGS={
  find:['[*] resolving target','[*] querying index','[*] cross-referencing','[+] fingerprinting matches'],
  month:['[*] sweeping calendar','[*] masking by month','[+] collecting hits'],
  list:['[*] tokenizing payload','[*] fuzzy matching','[!] resolving collisions','[+] correlating birthdays'],
  all:['[*] opening database','[*] streaming rows'],
  stats:['[*] profiling ages','[*] counting campuses','[*] computing upcoming','[+] building report']
};
function rhex(n){let s='';for(let i=0;i<n;i++)s+='0123456789abcdef'[Math.floor(Math.random()*16)];return s;}
async function runScan(kind,work){
  const lines=SCAN_LOGS[kind]||SCAN_LOGS.find,bar=$('prog');
  bar.style.opacity=1;bar.style.width='8%';
  let i=0,pct=8;STATUS.className='';
  const tick=setInterval(()=>{
    STATUS.textContent=lines[i%lines.length]+'  0x'+rhex(6);i++;
    pct=Math.min(90,pct+(90-pct)*.2);bar.style.width=pct+'%';
  },190);
  const t0=Date.now();
  try{
    const r=await work();
    const w=Math.max(0,650-(Date.now()-t0));if(w)await sleep(w);
    return r;
  }finally{
    clearInterval(tick);bar.style.width='100%';
    setTimeout(()=>{bar.style.opacity=0;setTimeout(()=>{bar.style.width='0';},320);},160);
  }
}

/* components */
function meta(s){
  const a=[s['Student ID'],s['Birthday']].filter(Boolean).join('  /  ');
  const b=[s['Campus'],s['Program']].filter(Boolean).join('  /  ');
  return [a,b];
}
function item(s,opts){
  opts=opts||{};
  const it=el('div','item'+(opts.click?' click':''));
  const body=el('div','body');
  body.appendChild(el('div','nm',s['Name']||opts.fallback||'--'));
  const m=meta(s);
  if(m[0])body.appendChild(el('div','mt',m[0]));
  if(m[1])body.appendChild(el('div','mt',m[1]));
  const pills=el('div');
  if(s.DaysToBday===0)pills.appendChild(el('span','pill hot','birthday today'));
  if(opts.tag)pills.appendChild(el('span','pill mute',opts.tag));
  if(pills.childNodes.length)body.appendChild(pills);
  it.appendChild(body);
  const side=el('div','side');
  if(s.Age!==''&&s.Age!=null){const a=el('div','age',s.Age);a.appendChild(el('small',null,'yrs'));side.appendChild(a);}
  const mo=opts.month?MONTHS[opts.month].slice(0,3):'';
  const dt=[mo,opts.day].filter(Boolean).join(' ');
  if(dt)side.appendChild(el('div','dt',dt));
  if(side.childNodes.length)it.appendChild(side);
  if(opts.click)it.onclick=opts.click;
  return it;
}
function listOf(rows,fn,cap){
  const box=el('div','list');
  if(!rows.length){box.appendChild(el('div','empty','nothing here'));return box;}
  (cap&&rows.length>cap?rows.slice(0,cap):rows).forEach(r=>box.appendChild(fn(r)));
  if(cap&&rows.length>cap)box.appendChild(el('div','empty','showing '+cap+' of '+rows.length+', refine your search to see the rest'));
  return box;
}
function sum(parts){
  const s=el('div','sum');
  parts.filter(Boolean).forEach(p=>{const sp=el('span');sp.innerHTML=p;s.appendChild(sp);});
  return s;
}
function h3(t,n){const h=el('div','h3',t);if(n!==undefined&&n!==null)h.appendChild(el('em',null,n));return h;}
function btn(label,fn){const b=el('button','btn ghost sm',label);b.onclick=fn;return b;}
function nextTxt(s){
  const d=s.DaysToBday;
  if(d===''||d==null)return '--';
  if(d===0)return 'today';
  return 'in '+d+' day'+(d===1?'':'s');
}
function bars(pairs,onClick,wide){
  const box=el('div','bars'+(wide?' wide':''));
  const max=Math.max(1,...pairs.map(p=>p[1]));
  pairs.forEach(([k,v])=>{
    const r=el('div','brow');
    if(onClick){r.style.cursor='pointer';r.title='tap to list students';r.onclick=()=>onClick(k);}
    r.appendChild(el('span','bk',k));
    const t=el('span','bt'),f=el('i');f.style.width=(v?Math.max(3,v/max*100):0)+'%';t.appendChild(f);r.appendChild(t);
    r.appendChild(el('span','bv',v));
    box.appendChild(r);
  });
  return box;
}

/* renderers */
function renderDossier(s,back){
  OUT.innerHTML='';
  const nm=el('div','dossier-name');nm.appendChild(el('small',null,'target acquired'));
  nm.appendChild(document.createTextNode(s['Name']||'--'));OUT.appendChild(nm);
  const kv=el('div','kv');
  [['Student ID',s['Student ID']],['Birthday',s['Birthday']],
   ['Age',s.Age===''?'--':s.Age+' years old','accent'],['Zodiac',s.Zodiac],['Next bday',nextTxt(s),'accent'],
   ['Campus',s['Campus']],['Program',s['Program']]].forEach(([k,v,c])=>{
    const r=el('div');r.appendChild(el('span',null,k));r.appendChild(el('span',c||null,v===''||v==null?'--':v));kv.appendChild(r);
  });
  OUT.appendChild(kv);
  if(back)OUT.appendChild(btn('< back to candidates',()=>renderCandidates(back)));
  window.scrollTo({top:0,behavior:'smooth'});
}
function renderCandidates(j){
  OUT.innerHTML='';
  OUT.appendChild(sum(['<b>'+j.count+'</b> candidates for <b>'+esc(j.query)+'</b>',
    j.count>j.candidates.length?'showing <b>'+j.candidates.length+'</b>':'']));
  const box=el('div','list');
  j.candidates.forEach(s=>box.appendChild(item(s,{
    month:MONTHS.indexOf(s.Month),day:s.Day,
    tag:s.score<1?'match '+Math.round(s.score*100)+'%':null,
    click:()=>renderDossier(s,j)})));
  OUT.appendChild(box);
  OUT.appendChild(el('div','hint','tap a name to open'));
}
function renderMonth(d){
  OUT.innerHTML='';
  OUT.appendChild(sum(['<b>'+d.count+'</b> birthdays in <b>'+esc(d.monthName.toLowerCase())+'</b>']));
  const mi=MONTHS.indexOf(d.monthName);
  OUT.appendChild(listOf(d.rows,r=>item(r,{month:mi,day:r.day,tag:r.DaysToBday===''||r.DaysToBday===0?null:nextTxt(r)}),RENDER_CAP));
}
function renderAll(d){
  OUT.innerHTML='';
  OUT.appendChild(sum(['<b>'+d.count+'</b> students']));
  const f=el('input');f.placeholder='filter by name, id, campus, month, age';f.setAttribute('autocapitalize','off');
  OUT.appendChild(f);
  const holder=el('div');holder.style.marginTop='10px';OUT.appendChild(holder);
  const cols=['Student ID','Name','Birthday','Age','Campus','Program','Month','Day'];
  const draw=q=>{
    q=(q||'').trim().toLowerCase();
    const rows=q?d.rows.filter(r=>cols.some(c=>String(r[c]).toLowerCase().includes(q))):d.rows;
    holder.innerHTML='';
    holder.appendChild(sum(['showing <b>'+Math.min(rows.length,RENDER_CAP)+'</b> of <b>'+rows.length+'</b>']));
    holder.appendChild(listOf(rows,r=>item(r,{month:MONTHS.indexOf(r.Month),day:r.Day}),RENDER_CAP));
  };
  f.oninput=()=>draw(f.value);draw('');
}
function renderList(a){
  OUT.innerHTML='';
  const p=['<b>'+a.total+'</b> lines','<b>'+a.matchedCount+'</b> matched','<b>'+a.thisMonthCount+'</b> in <b>'+esc((a.monthName||'--').toLowerCase())+'</b>'];
  if(a.duplicates&&a.duplicates.length)p.push('<b>'+a.duplicates.length+'</b> duplicates skipped');
  if(a.unmatchedCount)p.push('<b>'+a.unmatchedCount+'</b> unmatched');
  OUT.appendChild(sum(p));
  if(a.lastError)OUT.appendChild(el('div','err','server error: '+a.lastError));
  if(!a.matchedCount&&!a.unmatchedCount)return;

  const mk=e=>item(e.student||{},{fallback:e.line,month:e.birthdayMonth,day:e.birthdayDay,
    tag:e.matchType==='inline'?'date from line':(e.matched&&e.matchType!=='id'&&e.matchType!=='exact'?e.matchType+' '+e.score:null)});

  OUT.appendChild(h3('birthday in '+(a.monthName||'--').toLowerCase(),a.thisMonth.length));
  OUT.appendChild(listOf(a.thisMonth,mk,RENDER_CAP));
  if(a.matchedOther.length){
    OUT.appendChild(h3('other months',a.matchedOther.length));
    OUT.appendChild(listOf(a.matchedOther,mk,RENDER_CAP));
  }
  if(a.unmatched.length){
    OUT.appendChild(h3('unmatched',a.unmatched.length));
    OUT.appendChild(listOf(a.unmatched,u=>{
      const it=el('div','item'),b=el('div','body');
      b.appendChild(el('div','nm',u.line));b.appendChild(el('div','mt',u.error||'no match'));
      it.appendChild(b);return it;
    },RENDER_CAP));
  }
}
function renderStats(d){
  OUT.innerHTML='';
  const grid=el('div','cards');
  [['students',d.total],['dated',d.withBirthday],['avg age',d.avgAge==null?'--':d.avgAge],['today',d.today.length]].forEach(([k,v])=>{
    const c=el('div','card');c.appendChild(el('b',null,v));c.appendChild(el('small',null,k));grid.appendChild(c);
  });
  OUT.appendChild(grid);
  if(d.today.length){
    OUT.appendChild(h3('birthday today',d.today.length));
    OUT.appendChild(listOf(d.today,r=>item(r,{month:MONTHS.indexOf(r.Month),day:r.Day}),RENDER_CAP));
  }
  OUT.appendChild(h3('incoming, next '+d.days+' days',d.upcomingCount));
  OUT.appendChild(listOf(d.upcoming,r=>item(r,{month:MONTHS.indexOf(r.Month),day:r.Day,tag:nextTxt(r)}),RENDER_CAP));
  if(d.byCampus.length){
    OUT.appendChild(h3('students per campus',d.byCampus.length+' campuses'));
    OUT.appendChild(bars(d.byCampus,loadCampus,true));
    OUT.appendChild(el('div','hint','tap a campus to list its students'+(d.noCampus?'  |  '+d.noCampus+' with no campus':'')));
  }
  OUT.appendChild(h3('by month'));
  OUT.appendChild(bars(MONTHS.slice(1).map((m,i)=>[m.slice(0,3).toLowerCase(),d.byMonth[i+1]||0])));
  if(d.byProgram.length){OUT.appendChild(h3('by program'));OUT.appendChild(bars(d.byProgram,null,true));}
  if(d.byAge.length){OUT.appendChild(h3('by age'));OUT.appendChild(bars(d.byAge.map(a=>['age '+a[0],a[1]])));}
}

/* network */
async function withTimeout(url,opts,ms){
  const c=new AbortController(),t=setTimeout(()=>c.abort(),ms);
  try{return await fetch(url,{...opts,signal:c.signal});}finally{clearTimeout(t);}
}
async function readErr(r){try{return (await r.json()).error||('HTTP '+r.status);}catch(_){return 'HTTP '+r.status;}}
const fail=e=>setStatus(e.name==='AbortError'?'[!] connection timed out':'[!] request failed: '+e.message,'bad');

async function lookupSingle(){
  const q=SINGLE.value.trim();
  if(!q){setStatus('[!] enter a name or student ID','bad');return;}
  SINGLE.blur();setBusy(true);
  try{
    const r=await runScan('find',()=>withTimeout('/search?q='+encodeURIComponent(q),{},15000));
    if(r.status===404){OUT.innerHTML='';setStatus('[-] no match for "'+q+'"','bad');return;}
    if(!r.ok){setStatus('[!] '+(await readErr(r)),'bad');return;}
    const j=await r.json();
    setStatus('[+] '+j.count+' found in '+j.elapsedMs+' ms','ok');
    if(j.candidates.length===1)renderDossier(j.candidates[0],null);else renderCandidates(j);
  }catch(e){fail(e);}finally{setBusy(false);checkServer();}
}
async function scanWorkbook(){
  setBusy(true);
  try{
    const r=await runScan('month',()=>withTimeout('/birthdays?month='+monthSel(),{},15000));
    if(!r.ok){setStatus('[!] '+(await readErr(r)),'bad');return;}
    const j=await r.json();setStatus('[+] done in '+j.elapsedMs+' ms','ok');renderMonth(j);
  }catch(e){fail(e);}finally{setBusy(false);checkServer();}
}
async function scanAll(){
  setBusy(true);
  try{
    const r=await runScan('all',()=>withTimeout('/all',{},30000));
    if(!r.ok){setStatus('[!] '+(await readErr(r)),'bad');return;}
    const j=await r.json();setStatus('[+] done in '+j.elapsedMs+' ms','ok');renderAll(j);
  }catch(e){fail(e);}finally{setBusy(false);checkServer();}
}
async function loadStats(){
  setBusy(true);
  try{
    const r=await runScan('stats',()=>withTimeout('/stats?days='+$('days').value,{},20000));
    if(!r.ok){setStatus('[!] '+(await readErr(r)),'bad');return;}
    const j=await r.json();STATS_LOADED=true;
    setStatus('[+] recon done in '+j.elapsedMs+' ms','ok');renderStats(j);
  }catch(e){fail(e);}finally{setBusy(false);checkServer();}
}
async function loadCampus(name){
  setBusy(true);
  try{
    const r=await runScan('all',()=>withTimeout('/all?campus='+encodeURIComponent(name),{},30000));
    if(!r.ok){setStatus('[!] '+(await readErr(r)),'bad');return;}
    const j=await r.json();
    setStatus('[+] '+j.count+' students in '+name,'ok');renderAll(j);
    OUT.insertBefore(btn('< back to intel',()=>loadStats()),OUT.firstChild);
    window.scrollTo({top:0,behavior:'smooth'});
  }catch(e){fail(e);}finally{setBusy(false);checkServer();}
}
async function batchCheck(raw,month){
  const lines=raw.split('\n').map(s=>s.trim()).filter(Boolean);
  const batches=[];for(let i=0;i<lines.length;i+=BATCH_SIZE)batches.push(lines.slice(i,i+BATCH_SIZE));
  const agg={monthName:null,total:0,duplicates:[],thisMonth:[],matchedOther:[],unmatched:[],failed:0,lastError:null};
  for(let b=0;b<batches.length;b++){
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
  return agg;
}
async function checkList(){
  const raw=PASTE.value,file=$('file').files[0];
  if(!raw.trim()&&!file){setStatus('[!] paste names or attach a txt file','bad');return;}
  const month=monthSel(),t0=performance.now();
  setBusy(true);
  try{
    let agg;
    if(file){
      agg=await runScan('list',async()=>{
        const fd=new FormData();fd.append('file',file);fd.append('month',String(month));
        const r=await withTimeout('/birthdays/list',{method:'POST',body:fd},120000);
        if(!r.ok)throw new Error(await readErr(r));
        return await r.json();
      });
    }else{
      agg=await runScan('list',()=>batchCheck(raw,month));
    }
    setStatus('[+] done in '+Math.round(performance.now()-t0)+' ms'+(agg.failed?'  |  '+agg.failed+' batch(es) failed':''),agg.failed?'bad':'ok');
    renderList(agg);
  }catch(e){fail(e);}finally{setBusy(false);checkServer();}
}
function clearAll(){
  PASTE.value='';safe(()=>localStorage.removeItem(LS_PASTE));
  $('file').value='';$('fileLbl').textContent='attach .txt';
  SINGLE.value='';OUT.innerHTML='';setStatus('');
}

$('singleBtn').onclick=lookupSingle;
$('scanBtn').onclick=scanWorkbook;
$('allBtn').onclick=scanAll;
$('checkBtn').onclick=checkList;
$('clearBtn').onclick=clearAll;
$('statsBtn').onclick=loadStats;
$('days').onchange=()=>{if(STATS_LOADED)loadStats();};
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
        "today": _today().isoformat(),
        "tz_offset_hours": TZ_OFFSET_HOURS,
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


@app.get("/search")
def search():
    """Candidate search. 'james' -> every student with James in their name.
    An exact student ID returns just that one student."""
    t0 = time.time()
    try:
        q = (request.args.get("q") or "").strip()
        if not q:
            return jsonify({"error": "q required"}), 400
        idx = get_index()
        m = match_index(idx, q)
        mode = "name"
        if m is not None and m["matchType"] in ("id", "id-token"):
            found, total, mode = [(1.0, m["index"])], 1, "id"
        else:
            found, total = find_candidates(idx, q)
            if not found and m is not None:
                found, total = [(m["score"], m["index"])], 1
        if not found:
            return jsonify({"found": False, "query": q, "count": 0, "candidates": [],
                            "elapsedMs": int((time.time() - t0) * 1000)}), 404
        cands = []
        for sc, i in found:
            s = full_student(idx, i)
            s["score"] = round(sc, 3)
            cands.append(s)
        return jsonify({
            "found": True, "query": q, "mode": mode,
            "count": total, "candidates": cands,
            "elapsedMs": int((time.time() - t0) * 1000),
        }), 200
    except Exception as e:
        return _err("/search", e, t0)


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
        today = _today()
        out = []
        for row in idx["rows"]:
            if row["birthday_month"] != m_num:
                continue
            item = row_to_public(row, today)
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
        today = _today()
        out = []
        for row in idx["rows"]:
            item = row_to_public(row, today)
            item["Month"] = _MONTH_NAMES[row["birthday_month"]] if row["birthday_month"] else ""
            item["Day"] = row["birthday_day"] if row["birthday_day"] is not None else ""
            item["_row"] = row["row"]
            out.append(item)

        q = (request.args.get("q") or "").strip().lower()
        if q:
            fields = VISIBLE + ["Month", "Age"]
            out = [r for r in out if any(q in str(r.get(f, "")).lower() for f in fields)]

        c_raw = (request.args.get("campus") or "").strip().lower()
        if c_raw:
            out = [r for r in out if r["Campus"].strip().lower() == c_raw]

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


@app.get("/stats")
def stats():
    """Database recon: totals, average age, today's birthdays, incoming birthdays, distributions."""
    t0 = time.time()
    try:
        try:
            days = max(1, min(365, int(request.args.get("days", "30"))))
        except ValueError:
            days = 30
        idx = get_index()
        today = _today()
        by_month = [0] * 13
        campus, program, ages_c = Counter(), Counter(), Counter()
        ages = []
        today_list, upcoming = [], []
        with_bday = 0

        for row in idx["rows"]:
            m = row["birthday_month"]
            if m:
                by_month[m] += 1
                with_bday += 1
            if row["campus"].strip():
                campus[row["campus"].strip()] += 1
            if row["program"].strip():
                program[row["program"].strip()] += 1

            pub = row_to_public(row, today)
            if pub["Age"] != "":
                ages.append(pub["Age"])
                ages_c[pub["Age"]] += 1
            dd = pub["DaysToBday"]
            if dd != "" and dd <= days:
                pub["Month"] = _MONTH_NAMES[m] if m else ""
                pub["Day"] = row["birthday_day"]
                (today_list if dd == 0 else upcoming).append(pub)

        upcoming.sort(key=lambda x: (x["DaysToBday"], x["Name"]))
        today_list.sort(key=lambda x: x["Name"])

        return jsonify({
            "total": len(idx["rows"]),
            "withBirthday": with_bday,
            "avgAge": round(sum(ages) / len(ages), 1) if ages else None,
            "minAge": min(ages) if ages else None,
            "maxAge": max(ages) if ages else None,
            "today": today_list,
            "days": days,
            "upcomingCount": len(upcoming),
            "upcoming": upcoming[:100],
            "byMonth": by_month,
            "byCampus": campus.most_common(),
            "noCampus": sum(1 for r in idx["rows"] if not r["campus"].strip()),
            "byProgram": program.most_common(8),
            "byAge": sorted(ages_c.items()),
            "asOf": today.isoformat(),
            "elapsedMs": int((time.time() - t0) * 1000),
        }), 200
    except Exception as e:
        return _err("/stats", e, t0)


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
        today = _today()
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
                dv = _derived(inline_bd.month, inline_bd.day,
                              inline_bd.year if inline_bd.year != 1900 else None, today)
                entry.update({
                    "matched": False, "rowNumber": None, "score": None, "matchType": "inline",
                    "student": {"Student ID": "", "Name": name_part or line,
                                "Birthday": inline_bd.strftime("%B %d, %Y"),
                                "Campus": "", "Program": "",
                                "Age": dv["Age"], "Zodiac": dv["Zodiac"],
                                "DaysToBday": dv["DaysToBday"]},
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

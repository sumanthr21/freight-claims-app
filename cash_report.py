"""
Engine for the Test Equity Cash Application Specialist report.

What it does
------------
1. Reads the agents' EOD workbooks (Varsha's and Gokulnath's layouts).
2. Appends their rows to the data sheets of the master report
   (Billtrust, Keyoff, Unapplied Cash, Credit Card, and - if present - COA,
   Unposted Report, Correlation HISCO), filling Month / Date / Week / Name / Task.
3. Sets the "Daily - Report" date cell on the current month sheet (e.g. Nov'26!B6).
4. Marks the workbook "recalculate on open" so Daily / Weekly / Monthly / month
   sheets refresh the first time the file is opened in Excel.

The report is edited at the XML level (not with openpyxl) so that charts,
comments, formatting and formulas are kept exactly as they are.
Only the standard library + openpyxl (for reading the small EOD files) are used.
"""
import datetime as dt
import difflib
import io
import re
import zipfile
from collections import Counter
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
EPOCH = dt.date(1899, 12, 30)

# --------------------------------------------------------------------------
# Settings you may want to edit later
# --------------------------------------------------------------------------
# A name in the report must appear at least this many times to count as an agent.
MIN_ROWS_FOR_AGENT = 20

# Spellings the report already uses. Anything the agents type that is close to
# one of these is corrected; anything unknown is kept as typed and flagged.
CANON = {
    ("Billtrust", "task type"): ["Match Review", "Correlation Review", "Correlation Assist"],
    ("Billtrust", "status"): ["Completed", "Pending", "In Progress"],
    ("Keyoff", "task type"): ["Key off"],
    ("Keyoff", "status"): ["Completed", "Pending"],
    ("Unapplied Cash", "status"): ["Completed", "In Progress", "Pending"],
}
ALIASES = {"complete": "Completed", "in-progress": "In Progress", "inprogress": "In Progress"}

# EOD sheet name -> where the rows go.
#   fields: report column (header, occurrence)  <-  EOD column (header, occurrence)
#   consts: fixed values written to the report column
#   letter-targets ("@O") are used where the report column has no header
ROUTES = {
    "billtrust": dict(
        target="Billtrust", task="Billtrust Review",
        fields=[("customer name", 0, "customer name", 0), ("invoice number", 0, "invoice number", 0),
                ("amount", 0, "amount", 0), ("task type", 0, "task type", 0),
                ("status", 0, "status", 0), ("comment", 0, "comment", 0)]),
    "tematchreview": dict(
        target="Billtrust", task="Billtrust Review",
        fields=[("customer name", 0, "customer name", 0), ("invoice number", 0, "invoice", 0),
                ("amount", 0, "amount", 0)],
        consts={("task type", 0): "Match Review", ("status", 0): "Completed"},
        tick_col="match review"),
    "keyoffs": dict(
        target="Keyoff", task="Keyoff",
        fields=[("customer name", 0, "customer name", 0), ("account number", 0, "account number", 0),
                ("invoice number", 0, "invoice number", 0), ("amount", 0, "amount", 0),
                ("task type", 0, "task type", 0), ("status", 0, "status", 0),
                ("comments", 0, "comment", 0)]),
    "unappliedcash": dict(
        target="Unapplied Cash", task="Unapplied Cash",
        fields=[("invoice date", 0, "invoice date", 0), ("journal", 0, "journal", 0),
                ("amount", 0, "amount", 0), ("reference", 0, "reference", 0),
                ("sequence", 0, "sequence", 0), ("status", 0, "status", 0),
                ("journal", 1, "journal", 1), ("agent", 0, "agent", 0),
                ("@O", 0, "notes", 0)]),
    "ccpte": dict(
        target="Credit Card", task="Credit Card Processing (Test Equity)",
        fields=[("customer name", 0, "customer name", 0), ("customer number", 0, "customer number", 0),
                ("invoice", 0, "invoice", 0), ("amount", 0, "amount", 0),
                ("journal", 0, "journal off set", 0), ("journal oeet", 0, "journal oeet", 0)]),
    "ccphisco": dict(
        target="Credit Card", task="Credit Card Processing (Hisco)",
        fields=[("customer name", 0, "customer name", 0), ("customer number", 0, "customer number", 0),
                ("invoice", 0, "invoice", 0), ("amount", 0, "amount", 0),
                ("journal", 0, "journal off set", 0), ("journal oeet", 0, "journal oeet", 0)]),
}
# Sheets with no sample EOD yet: matched to the report by column names (generic).
GENERIC_ROUTES = {
    "coa": ("COA", "COA (Cash On Accounts)"),
    "unpostedreport": ("Unposted Report", "Unposted Report"),
    "correlationhisco": ("Correlation HISCO", "Correlation HISCO"),
}
AGENT_HEADERS = ["agent name", "employee name", "agent", "specialist", "name"]
SKIP_HEADERS = {"sno", "sn", "s no", "sl no", "no"}


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------
def norm(s):
    """Lower-case, no punctuation, single spaces: 'Customer  Name ' -> 'customer name'."""
    if s is None:
        return ""
    s = str(s).replace("\xa0", " ").lower()
    s = re.sub(r"[^a-z0-9@ ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def key(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def clean_text(v):
    if v is None:
        return None
    if isinstance(v, str):
        v = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", v.replace("\xa0", " ")).strip()
        return v or None
    return v


def col_letter(n):  # 1 -> A
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def serial(d):
    return (d - EPOCH).days


def to_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if 30000 < v < 80000:
            return EPOCH + dt.timedelta(days=int(v))
        return None
    s = str(v).replace("\xa0", " ").strip()
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d-%m-%Y", "%d-%b-%Y", "%d-%b-%y"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def money(v):
    """'$   (8,763.86)' -> -8763.86 ; '$2,823.58 USD' -> 2823.58 ; 306.18 -> 306.18.
    Returns (value, ok)."""
    if v is None:
        return None, True
    if isinstance(v, bool):
        return None, False
    if isinstance(v, (int, float)):
        return (int(v) if float(v).is_integer() else float(v)), True
    s = str(v).replace("\xa0", " ").strip()
    if not s:
        return None, True
    neg = ("(" in s and ")" in s) or s.startswith("-") or s.endswith("-") or "- " in s[:3]
    num = re.sub(r"[^0-9.]", "", s.replace(",", ""))
    if not num or num.count(".") > 1:
        return s, False
    x = float(num)
    if neg:
        x = -x
    return (int(x) if x.is_integer() and "." not in num else x), True


def tidy(v):
    """Keep numbers as numbers (12.0 -> 12), strings trimmed."""
    v = clean_text(v)
    if isinstance(v, float) and v.is_integer() and abs(v) < 1e15:
        return int(v)
    return v


def week_of(d):
    """Mon-Sun weeks. Week 1 = week containing the month's first weekday.
    (Matches 100% of the 58,887 rows already in the report.)"""
    f = d.replace(day=1)
    while f.weekday() >= 5:
        f += dt.timedelta(days=1)
    monday0 = f - dt.timedelta(days=f.weekday())
    return f"Week {(d - monday0).days // 7 + 1}"


FULL = ["January", "February", "March", "April", "May", "June", "July", "August",
        "September", "October", "November", "December"]


def month_variants(d):
    yy = d.strftime("%y")
    full = FULL[d.month - 1]
    return [f"{full[:3]}'{yy}", f"{full}'{yy}", f"{full[:4]}'{yy}", f"Sept'{yy}" if d.month == 9 else f"{full[:3]}'{yy}"]


# --------------------------------------------------------------------------
# Reading the EOD files
# --------------------------------------------------------------------------
def parse_eod(filename, data):
    """Return dict(file, sheets=[...], warnings=[...]) for one EOD workbook.
    Each sheet entry: name, route, target, task, rows=[{date, agent, fields}], notes."""
    import openpyxl
    out = dict(file=filename, sheets=[], warnings=[])
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    except Exception as e:  # noqa
        out["warnings"].append(f"{filename}: cannot open this file ({e}).")
        return out
    for ws in wb.worksheets:
        k = key(ws.title)
        if k in ROUTES:
            route = dict(ROUTES[k], generic=False)
        elif k in GENERIC_ROUTES:
            route = dict(target=GENERIC_ROUTES[k][0], task=GENERIC_ROUTES[k][1], generic=True)
        else:
            if ws.max_row > 1 and any(c is not None for row in ws.iter_rows(max_row=3, values_only=True) for c in row):
                out["warnings"].append(f"{filename}: sheet '{ws.title.strip()}' is not recognised and was skipped.")
            continue
        sh = _parse_sheet(ws, route, filename)
        out["sheets"].append(sh)
    if not out["sheets"]:
        out["warnings"].append(f"{filename}: no known EOD sheets found "
                               "(expected e.g. Billtrust, Key offs, Unapplied cash, CCP_TE, CCP_HISCO, TE_ MATCH REVIEW).")
    return out


def _parse_sheet(ws, route, filename):
    rows = list(ws.iter_rows(values_only=True))
    sh = dict(name=ws.title.strip(), file=filename, target=route["target"], task=route["task"],
              rows=[], notes=[], generic=route["generic"])
    # header row = first row with >= 3 text cells
    hi = next((i for i, r in enumerate(rows[:8]) if sum(1 for c in r if isinstance(c, str) and c.strip()) >= 3), None)
    if hi is None:
        sh["notes"].append(("error", f"{filename} / {sh['name']}: could not find the header row."))
        return sh
    hdr = rows[hi]
    pos = {}  # (norm header, occurrence) -> index
    seen = Counter()
    for i, h in enumerate(hdr):
        n = norm(h)
        if n:
            pos[(n, seen[n])] = i
            seen[n] += 1

    def get(r, h, occ=0):
        i = pos.get((h, occ))
        return r[i] if i is not None and i < len(r) else None

    # which EOD column holds the date / agent
    date_i = pos.get(("date", 0))
    agent_key = next((a for a in AGENT_HEADERS if (a, 0) in pos), None)

    if route["generic"]:
        fields, unmatched = _generic_fields(hdr, route["target"], pos, agent_key)
        if unmatched:
            sh["notes"].append(("warn", f"{filename} / {sh['name']}: columns not copied (no matching report column): "
                                        + ", ".join(unmatched)))
        sh["notes"].append(("warn", f"{sh['name']}: loaded by matching column names only - no sample EOD was "
                                    "used to test this sheet, so check the rows after the first run."))
        route = dict(route, fields=fields)
    else:
        missing = [e for (_, _, e, eo) in route["fields"] if (e, eo) not in pos]
        if missing:
            sh["notes"].append(("warn", f"{filename} / {sh['name']}: expected column(s) not found: "
                                        + ", ".join(sorted(set(missing))) + " - those cells are left blank."))
    last_date = last_agent = None
    for r in rows[hi + 1:]:
        vals = {}
        has = False
        for (th, to, eh, eo) in route["fields"]:
            v = get(r, eh, eo)
            v = clean_text(v)
            if v is not None:
                has = True
            vals[(th, to)] = v
        for ck in (route.get("consts") or {}):
            vals[ck] = route["consts"][ck]
        if not has:
            continue
        tick = route.get("tick_col")
        if tick is not None:
            tv = get(r, tick)
            if not (tv is True or (isinstance(tv, str) and tv.strip().lower() in ("true", "yes", "y", "1")) or tv == 1):
                sh.setdefault("skipped_unticked", 0)
                sh["skipped_unticked"] += 1
                continue
        d = to_date(r[date_i]) if date_i is not None and date_i < len(r) else None
        if d is None:
            d = last_date
        else:
            last_date = d
        ag = clean_text(get(r, agent_key)) if agent_key else None
        if ag is None:
            ag = last_agent
        else:
            last_agent = ag
        sh["rows"].append(dict(date=d, agent=ag, vals=vals))
    if sh.get("skipped_unticked"):
        sh["notes"].append(("warn", f"{filename} / {sh['name']}: {sh['skipped_unticked']} row(s) without a TRUE tick in "
                                    "'Match Review' were NOT loaded."))
    return sh


def _generic_fields(hdr, target, pos, agent_key):
    skip = set(SKIP_HEADERS) | {"date"} | set(AGENT_HEADERS)
    fields, unmatched = [], []
    seen = Counter()
    for h in hdr:
        n = norm(h)
        if not n:
            continue
        occ = seen[n]
        seen[n] += 1
        if n in skip:
            continue
        fields.append((n, occ, n, occ))   # report header == EOD header (checked later against the report)
    return fields, unmatched


# --------------------------------------------------------------------------
# Reading the report (zip / XML)
# --------------------------------------------------------------------------
def _read_workbook(zf):
    wb = zf.read("xl/workbook.xml").decode("utf8")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf8")
    rid = {}
    for m in re.finditer(r"<Relationship\b([^>]*)/?>", rels):
        a = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1)))
        rid[a["Id"]] = a["Target"].lstrip("/")
    sheets = {}
    for m in re.finditer(r"<sheet\b([^>]*)/?>", wb):
        a = dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(1)))
        t = rid[a["r:id"]]
        sheets[a["name"].replace("&apos;", "'").replace("&amp;", "&")] = t if t.startswith("xl/") else "xl/" + t
    return sheets


def _shared_strings(zf):
    out = []
    if "xl/sharedStrings.xml" not in zf.namelist():
        return out
    with zf.open("xl/sharedStrings.xml") as f:
        for _, el in ET.iterparse(f, events=("end",)):
            if el.tag == NS + "si":
                out.append("".join(t.text or "" for t in el.iter(NS + "t")))
                el.clear()
    return out


def _cell_value(c, shared):
    t = c.get("t")
    if t == "inlineStr":
        return "".join(x.text or "" for x in c.iter(NS + "t"))
    v = c.find(NS + "v")
    if v is None or v.text is None:
        return None
    if t == "s":
        return shared[int(v.text)]
    if t in ("str", "e"):
        return v.text
    if t == "b":
        return v.text == "1"
    try:
        return float(v.text)
    except ValueError:
        return v.text


def scan_sheet(zf, path, shared):
    """Return header {letter: text}, set of (date_serial, name, task), Counter of names."""
    header, triples, names = {}, set(), Counter()
    cols = {}
    with zf.open(path) as f:
        for _, el in ET.iterparse(f, events=("end",)):
            if el.tag != NS + "row":
                continue
            rn = int(el.get("r"))
            if rn == 1:
                for c in el:
                    letters = re.match(r"[A-Z]+", c.get("r")).group(0)
                    v = _cell_value(c, shared)
                    if v not in (None, ""):
                        header[letters] = str(v)
                for L, h in header.items():
                    n = norm(h)
                    if n in ("date", "name", "task") and n not in cols:
                        cols[n] = L
            elif cols:
                got = {}
                for c in el:
                    L = re.match(r"[A-Z]+", c.get("r")).group(0)
                    if L in cols.values():
                        got[L] = _cell_value(c, shared)
                d, nm, tk = (got.get(cols.get("date")), got.get(cols.get("name")), got.get(cols.get("task")))
                if isinstance(d, float) and nm and tk:
                    nm = str(nm).strip()
                    triples.add((int(d), nm, str(tk).strip()))
                    names[nm] += 1
            el.clear()
    return header, triples, names


def _date_style_ids(styles_xml):
    s = styles_xml.decode("utf8")
    custom = {}
    for m in re.finditer(r'<numFmt\b[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"', s):
        custom[int(m.group(1))] = m.group(2)
    block = re.search(r"<cellXfs\b.*?</cellXfs>", s, re.S)
    ids = set()
    if not block:
        return ids
    for i, m in enumerate(re.finditer(r"<xf\b[^>]*>", block.group(0))):
        mf = re.search(r'numFmtId="(\d+)"', m.group(0))
        nf = int(mf.group(1)) if mf else 0
        if 14 <= nf <= 22 or 45 <= nf <= 47:
            ids.add(i)
        elif nf in custom:
            code = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", custom[nf]).lower()
            if re.search(r"[dmy]", code) and "general" not in code and "0.0" not in code and "#" not in code:
                ids.add(i)
    return ids


def _tail_info(sx):
    """Find the last row that holds data. Returns (last_row_number, styles{letter: s}, cut_start, cut_end)."""
    end = sx.rfind(b"</sheetData>")
    pos = sx.rfind(b"<row ", 0, end)
    nxt = end
    while pos >= 0:
        chunk = sx[pos:nxt]
        if b"<v>" in chunk or b"<is>" in chunk:
            break
        nxt = pos
        pos = sx.rfind(b"<row ", 0, pos)
    if pos < 0:
        return 1, {}, end, end
    rn = int(re.search(rb'r="(\d+)"', sx[pos:pos + 60]).group(1))
    styles = {}
    for m in re.finditer(rb'<c r="([A-Z]+)\d+"(?: s="(\d+)")?', sx[pos:nxt]):
        if m.group(2) is not None:
            styles[m.group(1).decode()] = int(m.group(2))
    return rn, styles, nxt, end


def _cell_xml(ref, v, style, date_ok_style):
    s = f' s="{style}"' if style is not None else ""
    if isinstance(v, bool):
        return f'<c r="{ref}"{s} t="b"><v>{int(v)}</v></c>'
    if isinstance(v, dt.date):
        st = style if style in date_ok_style[0] else date_ok_style[1]
        s = f' s="{st}"' if st is not None else ""
        return f'<c r="{ref}"{s}><v>{serial(v)}</v></c>'
    if isinstance(v, (int, float)):
        return f'<c r="{ref}"{s}><v>{repr(v) if isinstance(v, float) else v}</v></c>'
    return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{escape(str(v))}</t></is></c>'


# --------------------------------------------------------------------------
# Main update
# --------------------------------------------------------------------------
def _canon(sheet, field, value, notes):
    lst = CANON.get((sheet, field))
    if not lst or not isinstance(value, str):
        return value
    k = value.strip().lower()
    for c in lst:
        if k == c.lower():
            return c
    if k in ALIASES and ALIASES[k] in lst:
        return ALIASES[k]
    m = difflib.get_close_matches(value.strip(), lst, n=1, cutoff=0.82)
    if m:
        notes.add(("info", f"{sheet} {field}: '{value}' corrected to '{m[0]}'."))
        return m[0]
    notes.add(("warn", f"{sheet} {field}: new value '{value}' kept as typed (not one of {', '.join(lst)})."))
    return value


def preview(parsed_list, override_date=None):
    """Quick table for the app: one line per (file, sheet)."""
    lines = []
    for p in parsed_list:
        for sh in p["sheets"]:
            dates = sorted({r["date"] for r in sh["rows"] if r["date"]})
            agents = sorted({r["agent"] for r in sh["rows"] if r["agent"]})
            lines.append(dict(File=p["file"], Sheet=sh["name"], Goes_to=f"{sh['target']} / {sh['task']}",
                              Rows=len(sh["rows"]),
                              Date=", ".join(d.strftime("%d-%b-%Y") for d in dates) or "(none)",
                              Agent=", ".join(agents) or "(none)"))
    return lines


def detect_date(parsed_list):
    c = Counter(r["date"] for p in parsed_list for sh in p["sheets"] for r in sh["rows"] if r["date"])
    return c.most_common(1)[0][0] if c else None


def update_report(master, parsed_list, override_date=None):
    """master: bytes of the report. Returns (new_bytes or None, messages[(level,text)], stats dict)."""
    msgs, notes = [], set()
    for p in parsed_list:
        for w in p["warnings"]:
            msgs.append(("warn", w))
        for sh in p["sheets"]:
            for n in sh["notes"]:
                notes.add(n)

    sheets_in = [sh for p in parsed_list for sh in p["sheets"] if sh["rows"]]
    if not sheets_in:
        return None, msgs + [("error", "No rows to load. Nothing was changed.")], {}

    zin = zipfile.ZipFile(io.BytesIO(master))
    paths = _read_workbook(zin)
    shared = _shared_strings(zin)
    date_styles = _date_style_ids(zin.read("xl/styles.xml"))

    # ---- scan the target sheets -------------------------------------------------
    scans = {}
    for t in sorted({sh["target"] for sh in sheets_in}):
        if t not in paths:
            return None, msgs + [("error", f"The report has no sheet called '{t}'. Nothing was changed.")], {}
        scans[t] = scan_sheet(zin, paths[t], shared)
    known = Counter()
    for t, (_, _, names) in scans.items():
        known.update(names)
    agents = {k_.lower(): k_ for k_, n in known.items() if n >= MIN_ROWS_FOR_AGENT}

    def agent_name(raw):
        if not raw:
            return None
        toks = re.findall(r"[A-Za-z]+", str(raw).lower())
        for tk in toks[:2]:
            if tk in agents:
                return agents[tk]
        for tk in toks[:1]:
            for a in agents:
                if a.startswith(tk) or tk.startswith(a):
                    return agents[a]
        notes.add(("warn", f"Unknown agent name '{str(raw).strip()}' - kept as typed. "
                           f"Known agents: {', '.join(sorted(agents.values()))}."))
        return str(raw).strip()

    # month labels already used in the report (e.g. April'26, July'26, Nov'26)
    # -> read from column A of the first scanned sheet is not stored; use the month sheet names instead
    month_label_cache = {}

    def month_label(d):
        if (d.year, d.month) in month_label_cache:
            return month_label_cache[(d.year, d.month)]
        lab = None
        for v in month_variants(d):
            if v in paths:
                lab = v
                break
        if lab is None:
            lab = f"{FULL[d.month - 1][:3]}'{d.strftime('%y')}"
        month_label_cache[(d.year, d.month)] = lab
        return lab

    # ---- build rows ------------------------------------------------------------
    by_target = {}
    stats = Counter()
    problems = []
    dups = set()
    for sh in sheets_in:
        header, triples, _ = scans[sh["target"]]
        hmap = {}
        occ = Counter()
        for L in sorted(header, key=col_index):
            n = norm(header[L])
            hmap[(n, occ[n])] = L
            occ[n] += 1
        need = [hmap.get((x, 0)) for x in ("month", "date", "week", "name", "task")]
        if None in need:
            return None, msgs + [("error", f"Sheet '{sh['target']}' does not have the expected Month/Date/Week/Name/Task headers.")], {}
        cMonth, cDate, cWeek, cName, cTask = need
        for r in sh["rows"]:
            d = override_date or r["date"]
            nm = agent_name(r["agent"])
            if d is None:
                problems.append(f"{sh['file']} / {sh['name']}: a row has no date. Pick a date in the app.")
                break
            if nm is None:
                problems.append(f"{sh['file']} / {sh['name']}: a row has no agent name.")
                break
            trip = (serial(d), nm, sh["task"])
            if trip in triples:
                dups.add(f"{d.strftime('%d-%b-%Y')} {nm} ({sh['target']}: {sh['task']})")
                continue
            cells = {cMonth: month_label(d), cDate: d, cWeek: week_of(d), cName: nm, cTask: sh["task"]}
            for (tkey, v) in r["vals"].items():
                th, to = tkey
                L = th[1:].upper() if th.startswith("@") else hmap.get((th, to))
                if L is None:
                    notes.add(("warn", f"{sh['target']}: no column '{th}' in the report - value not copied."))
                    continue
                if v is None:
                    continue
                if th == "amount":
                    v2, ok = money(v)
                    if not ok:
                        notes.add(("warn", f"{sh['target']}: amount '{v}' is not a number - kept as text."))
                    v = v2
                elif th.endswith("date"):
                    dd = to_date(v)
                    if dd is None:
                        notes.add(("warn", f"{sh['target']}: could not read date '{v}' in '{th}'."))
                    v = dd or v
                else:
                    v = tidy(v)
                    v = _canon(sh["target"], th, v, notes)
                if v is not None:
                    cells[L] = v
            by_target.setdefault(sh["target"], []).append(cells)
            stats[(sh["target"], sh["task"], nm, d)] += 1
    if problems:
        return None, msgs + [("error", p) for p in problems] + [("error", "Nothing was changed.")], {}
    if dups:
        lines = "; ".join(sorted(dups))
        return None, msgs + [("error", f"Already loaded (date / agent): {lines}. Nothing was changed. "
                                       "If this is a mistake, restore the earlier report first.")], {}

    # ---- write the zip ----------------------------------------------------------
    replaced = {}
    for t, rows in by_target.items():
        sx = zin.read(paths[t])
        rn, styles, cut_a, cut_b = _tail_info(sx)
        date_fallback = styles.get(next((L for L, h in scans[t][0].items() if norm(h) == "date"), "B"))
        chunks = []
        r = rn
        for cells in rows:
            r += 1
            parts = []
            for L in sorted(cells, key=col_index):
                parts.append(_cell_xml(f"{L}{r}", cells[L], styles.get(L), (date_styles, date_fallback)))
            chunks.append(f'<row r="{r}">{"".join(parts)}</row>'.encode("utf8"))
        sx = sx[:cut_a] + b"".join(chunks) + sx[cut_b:]
        maxc = max([col_index(L) for L in scans[t][0]] + [col_index(L) for c in rows for L in c])
        sx = re.sub(rb'<dimension ref="[A-Z]+\d+:[A-Z]+\d+"/>',
                    f'<dimension ref="A1:{col_letter(maxc)}{r}"/>'.encode(), sx, count=1)
        replaced[paths[t]] = sx
        msgs.append(("ok", f"{t}: +{len(rows)} rows"))

    # month sheet date
    all_dates = sorted({k_[3] for k_ in stats})
    last = all_dates[-1]
    if len(all_dates) > 1:
        msgs.append(("warn", "Rows for more than one date were loaded ("
                             + ", ".join(d.strftime('%d-%b') for d in all_dates)
                             + "). The Daily date cell is set to the latest: " + last.strftime("%d-%b-%Y") + "."))
    msheet = next((v for v in month_variants(last) if v in paths), None)
    if msheet is None:
        msgs.append(("warn", f"There is no '{month_label(last)}' sheet yet, so the Daily date cell was not set. "
                             f"Copy the previous month's sheet, rename it {month_label(last)}, then load again "
                             "(see the guide: 'New month')."))
    else:
        mx = zin.read(paths[msheet]).decode("utf8")
        idx = shared.index("Daily - Report") if "Daily - Report" in shared else None
        row = None
        if idx is not None:
            m = re.search(r'<c r="A(\d+)"[^>]*t="s"[^>]*><v>%d</v>' % idx, mx)
            row = m.group(1) if m else None
        if row is None:
            msgs.append(("warn", f"Could not find the 'Daily - Report' date cell on '{msheet}'. Set it by hand to "
                                 + last.strftime("%d-%b-%Y") + "."))
        else:
            pat = re.compile(r'(<c r="B%s"[^>]*?)(?: t="\w+")?([^>]*>)<v>[^<]*</v>' % row)
            if pat.search(mx):
                mx = pat.sub(lambda m: f"{m.group(1)}{m.group(2)}<v>{serial(last)}</v>", mx, count=1)
                replaced[paths[msheet]] = mx.encode("utf8")
                msgs.append(("ok", f"{msheet}: Daily date (B{row}) set to {last.strftime('%d-%b-%Y')}"))
            else:
                msgs.append(("warn", f"Could not set B{row} on '{msheet}'. Set it by hand to {last.strftime('%d-%b-%Y')}."))

    wbx = zin.read("xl/workbook.xml").decode("utf8")
    if "fullCalcOnLoad" not in wbx:
        if re.search(r"<calcPr\b", wbx):
            wbx = re.sub(r"<calcPr\b", '<calcPr fullCalcOnLoad="1"', wbx, count=1)
        else:
            wbx = wbx.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
    replaced["xl/workbook.xml"] = wbx.encode("utf8")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zout:
        for item in zin.infolist():
            data = replaced.get(item.filename)
            if data is None:
                data = zin.read(item.filename)
            zout.writestr(item, data, compress_type=zipfile.ZIP_DEFLATED)
    zin.close()

    # ---- summary messages -------------------------------------------------------
    per = Counter()
    for (t, task, nm, d), n in stats.items():
        per[(task, nm, d)] += n
    for (task, nm, d), n in sorted(per.items(), key=lambda x: (x[0][2], x[0][1], x[0][0])):
        msgs.append(("ok", f"{d.strftime('%d-%b')} - {nm} - {task}: {n} rows"))
    for lv, tx in sorted(notes, key=lambda x: (x[0] != "warn", x[1])):
        msgs.append((lv, tx))
    return buf.getvalue(), msgs, dict(by_target={k_: len(v) for k_, v in by_target.items()}, date=last)

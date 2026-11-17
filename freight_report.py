"""Append agent EOD files to the 'Freight Claims' data sheet of the Hisco Freight Claims report.
Edits the .xlsx XML directly so charts, pivots and formulas survive (openpyxl would drop them)."""
import re, zipfile, io, datetime as dt
from collections import Counter
from xml.sax.saxutils import escape
import pandas as pd

SHEET = "Freight Claims"
COLS = ["Year", "Month", "Date", "Week", "Task", "Agent", "Tracking Number", "PO # OR WT #", "Status"]
TASK = "Freight Claims"
ALIASES = {"sent email to requestor": "Sent email to Requester"}   # spelling variants -> report spelling
EPOCH = dt.datetime(1899, 12, 30)

def month_label(d):  # report's own naming: April'26, June'26, July'26, others 3 letters
    return {4: "April", 6: "June", 7: "July"}.get(d.month, d.strftime("%b")) + "'" + d.strftime("%y")

def week_of(d):  # Mon-Sun weeks; Week 1 = week containing the month's first weekday (matches existing data 100%)
    first = d.replace(day=1)
    while first.weekday() > 4: first += dt.timedelta(days=1)
    mon = lambda x: x - dt.timedelta(days=x.weekday())
    return f"Week {(mon(d) - mon(first)).days // 7 + 1}"

def learn(report_bytes):
    """Read the report once: canonical status spellings, agent names, and (date, agent) pairs already loaded."""
    d = pd.read_excel(io.BytesIO(report_bytes), sheet_name=SHEET)
    c = Counter(d.Status.dropna().astype(str)); best = {}
    for s, n in c.items():
        k = s.lower()
        if k not in best or n > c[best[k]]: best[k] = s
    agents = [a for a, n in d.Agent.value_counts().items() if n > 20]
    done = {(pd.Timestamp(a).normalize(), b) for a, b in zip(d.Date, d.Agent)}
    return {"status": {**best, **ALIASES}, "agents": agents, "done": done}

def clean(eod, ctx):
    """Tidy one agent's EOD sheet -> rows in report column order. Returns (df, warnings)."""
    w = []
    df = eod.iloc[:, :5].copy(); df.columns = ["Date", "Name", "Tracking Number", "PO # OR WT #", "Status"]  # ignores stray extra columns
    df["Date"] = pd.to_datetime(df["Date"]).ffill()
    df = df.dropna(subset=["Name", "Status"]).copy()
    def agent(n):
        m = [a for a in ctx["agents"] if a.split()[0].lower() == str(n).strip().split()[0].lower()]
        if not m: w.append(f"Unknown agent '{n}' - kept as typed")
        return m[0] if m else str(n).strip()
    df["Agent"] = df.Name.map(agent)
    df["Status"] = df.Status.astype(str).str.strip().map(lambda s: ctx["status"].get(s.lower(), s))
    for s in sorted(set(df.Status) - set(ctx["status"].values())): w.append(f"New status '{s}' (not in report yet)")
    def tn(v):
        if pd.isna(v): return None
        if isinstance(v, float) and v.is_integer(): v = int(v)
        s = str(v).strip()
        return int(s) if s.isdigit() and len(s) <= 15 and not s.startswith("0") else s
    df["Tracking Number"] = df["Tracking Number"].map(tn)
    df["PO # OR WT #"] = df["PO # OR WT #"].map(lambda v: None if pd.isna(v) else str(v).strip())
    df["Year"] = df.Date.dt.year; df["Month"] = df.Date.map(month_label); df["Week"] = df.Date.map(week_of); df["Task"] = TASK
    return df[COLS], w

def _cell(ref, v, style):
    s = f' s="{style}"' if style else ""
    if v is None or (isinstance(v, float) and pd.isna(v)) or v == "": return ""
    if isinstance(v, (dt.datetime, pd.Timestamp)): return f'<c r="{ref}"{s}><v>{(pd.Timestamp(v).to_pydatetime() - EPOCH).days}</v></c>'
    if isinstance(v, (int, float)) and not isinstance(v, bool): return f'<c r="{ref}"{s}><v>{v}</v></c>'
    return f'<c r="{ref}"{s} t="inlineStr"><is><t>{escape(str(v))}</t></is></c>'

def _sheet_paths(z):
    wb = z.read("xl/workbook.xml").decode(); rels = z.read("xl/_rels/workbook.xml.rels").decode()
    rid = {}
    for tag in re.findall(r"<Relationship [^>]*>", rels):
        rid[re.search(r'Id="([^"]+)"', tag).group(1)] = re.search(r'Target="([^"]+)"', tag).group(1)
    return {n.replace("&apos;", "'"): "xl/" + rid[r].lstrip("/").replace("xl/", "") for n, r in re.findall(r'<sheet [^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', wb)}

def update_report(report_bytes, eod_frames):
    """eod_frames: list of raw DataFrames (one per agent file). Returns (new_xlsx_bytes, log)."""
    ctx = learn(report_bytes); log = []; parts = []
    for e in eod_frames:
        df, w = clean(e, ctx); log += [f"WARNING {x}" for x in w]; parts.append(df)
    rows = pd.concat(parts, ignore_index=True).sort_values(["Date", "Agent"], kind="stable")
    seen = {(pd.Timestamp(a).normalize(), b) for a, b in zip(rows.Date, rows.Agent)} & ctx["done"]
    if seen: raise ValueError("Already loaded (date / agent): " + "; ".join(f"{d:%d-%b-%Y} {a}" for d, a in sorted(seen)) + ". Nothing was added.")
    zin = zipfile.ZipFile(io.BytesIO(report_bytes)); paths = _sheet_paths(zin); p = paths[SHEET]; xml = zin.read(p).decode()
    last_row = max(int(x) for x in re.findall(r'<row r="(\d+)"', xml))
    last = re.findall(rf'<row r="{last_row}".*?</row>', xml, flags=re.S)[0]
    styles = dict(re.findall(r'<c r="([A-Z]+)\d+" s="(\d+)"', last)); L = "ABCDEFGHI"
    xr = [f'<row r="{i}">' + "".join(_cell(f"{L[j]}{i}", v, styles.get(L[j])) for j, v in enumerate(rec)) + "</row>"
          for i, rec in enumerate(rows.itertuples(index=False), start=last_row + 1)]
    xml = xml.replace("</sheetData>", "".join(xr) + "</sheetData>")
    xml = re.sub(r'<dimension ref="A1:([A-Z]+)\d+"', lambda m: f'<dimension ref="A1:{m.group(1)}{last_row + len(rows)}"', xml, count=1)
    new = {p: xml}
    for ag, n in rows.Agent.value_counts().items(): log.append(f"{ag}: +{n} rows")
    latest = rows.Date.max(); m = month_label(latest)
    if m in paths:
        mp = paths[m]; x = zin.read(mp).decode()
        x, n = re.subn(r'(<c r="B4"[^>]*>)<v>\d+</v>', lambda k: f'{k.group(1)}<v>{(latest.to_pydatetime() - EPOCH).days}</v>', x, count=1)
        new[mp] = x; log.append(f"{m}!B4 set to {latest:%d-%b-%Y}" if n else f"WARNING could not set {m}!B4")
    else: log.append(f"WARNING there is no '{m}' sheet yet - copy last month's sheet, rename it, then upload again")
    wbx = zin.read("xl/workbook.xml").decode()
    wbx = re.sub(r"<calcPr[^>]*/>", '<calcPr calcId="191029" fullCalcOnLoad="1"/>', wbx) if "<calcPr" in wbx else wbx.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
    new["xl/workbook.xml"] = wbx
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in zin.infolist():
            if it.filename != "xl/calcChain.xml": zo.writestr(it, new.get(it.filename, zin.read(it.filename)))
    log.insert(0, f"{SHEET}: +{len(rows)} rows (now {last_row + len(rows)}) for {rows.Date.min():%d-%b-%Y}")
    return out.getvalue(), log

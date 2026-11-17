"""Append agent EOD files (Customer Past Dues) to the 'Data' sheet of the Test Equity OS Order Maintenance report.
The report is large (275k data rows, pivots, charts), so the file is edited at XML level and STREAMED:
nothing is loaded whole into memory, and charts/pivots/formulas survive (openpyxl would drop them)."""
import re, zipfile, io, datetime as dt
from xml.sax.saxutils import escape
import pandas as pd

DATA = "Data"
COLS = ["Month", "Date", "Week", "Name", "Task", "Order #", "Order Suffix", "Line #", "Customer/Vendor/Warehouse",
        "Buyer Name", "Disposition", "Comments", "ADDITIONAL COMMENTS", "Due Date"]
LET = "ABCDEFGHIJKLMN"
TASK = "Past Due"
EPOCH = dt.datetime(1899, 12, 30)
CH, KEEP = 8_000_000, 8192

def week_of(d):  # Mon-Sun weeks; Week 1 = week holding the month's first weekday
    first = d.replace(day=1)
    while first.weekday() > 4: first += dt.timedelta(days=1)
    mon = lambda x: x - dt.timedelta(days=x.weekday())
    return f"Week {(mon(d) - mon(first)).days // 7 + 1}"

def month_label(d, sheets):
    full = {4: "April", 6: "June", 7: "July"}.get(d.month, d.strftime("%b")); yy = "'" + d.strftime("%y")
    for c in (full + yy, d.strftime("%b") + yy, d.strftime("%B") + yy):
        if c in sheets: return c
    return full + yy

def _sst(z):
    x = z.read("xl/sharedStrings.xml").decode()
    return [re.sub(r"<[^>]+>", "", s) for s in re.findall(r"<si>(.*?)</si>", x, flags=re.S)]

def _paths(z):
    wb = z.read("xl/workbook.xml").decode(); rels = z.read("xl/_rels/workbook.xml.rels").decode()
    rid = {re.search(r'Id="([^"]+)"', t).group(1): re.search(r'Target="([^"]+)"', t).group(1) for t in re.findall(r"<Relationship [^>]*>", rels)}
    return {n.replace("&apos;", "'"): "xl/" + rid[r].lstrip("/").replace("xl/", "") for n, r in re.findall(r'<sheet [^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', wb)}

def _tail(z, path, n=6_000_000):
    buf = b""
    with z.open(path) as f:
        while True:
            b = f.read(CH)
            if not b: break
            buf = (buf + b)[-n:]
    return buf

def learn(report_bytes):
    z = zipfile.ZipFile(io.BytesIO(report_bytes)); paths = _paths(z); sst = _sst(z)
    daily = z.read(paths["Daily"]).decode()   # agent names = header row 2 of the Daily sheet
    r2 = re.search(r'<row r="2".*?</row>', daily, flags=re.S).group(0)
    agents = [sst[int(v)] for _, v in re.findall(r'<c r="([B-Z]+2)"[^>]*t="s"[^>]*><v>(\d+)</v>', r2)]
    tail = _tail(z, paths[DATA]).decode("utf-8", "ignore")
    rows = list(re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', tail, flags=re.S))[1:]
    done, last = set(), rows[-1].group(2)
    for m in rows:
        d = re.search(r'<c r="B\d+"[^>]*><v>(\d+)</v>', m.group(2)); n = re.search(r'<c r="D\d+"([^>]*)>(.*?)</c>', m.group(2), flags=re.S)
        if d and n:
            nm = sst[int(re.search(r"<v>(\d+)</v>", n.group(2)).group(1))] if 't="s"' in n.group(1) else re.sub(r"<[^>]+>", "", n.group(2))
            done.add((int(d.group(1)), nm))
    styles = dict(re.findall(r'<c r="([A-Z]+)\d+" s="(\d+)"', last))
    return {"agents": agents, "done": done, "styles": styles, "sheets": set(paths)}

def guess_agent(filename, agents):
    f = re.sub(r"[^a-z]", "", filename.lower())
    for a in agents:
        k = re.sub(r"[^a-z]", "", a.lower())
        if k and (k in f or a.split()[0].lower() in f): return a
    return None

def clean(eod_path_or_file, agent, date, mapping=None):
    """One agent's EOD (first sheet) -> rows in Data-sheet column order, plus warnings."""
    xl = pd.ExcelFile(eod_path_or_file); df = xl.parse(0); w = []
    if mapping is None:   # buyer code -> full name, from the helper sheet found in the EOD itself (A1 = B1 = 'Buyer Name')
        mapping = {}
        for s in xl.sheet_names[1:]:
            h = xl.parse(s, header=None, nrows=1)
            if h.shape[1] > 1 and str(h.iloc[0, 0]).strip() == "Buyer Name" == str(h.iloc[0, 1]).strip():
                t = xl.parse(s, usecols=[0, 1]).dropna(); mapping = {str(a).strip().lower(): str(b).strip() for a, b in zip(t.iloc[:, 0], t.iloc[:, 1])}; break
    df = df.iloc[:, :9].copy(); df.columns = ["sno", "order", "suffix", "line", "vendor", "buyer", "disp", "comm", "task"]
    df = df.dropna(subset=["order", "disp"]).copy()
    names = set(mapping.values())
    def buyer(b):
        b = str(b).strip(); n = mapping.get(b.lower(), b)
        if mapping and n not in names: w.append(f"Buyer '{n}' is not in the buyer list - kept as typed")
        return n
    df["buyer"] = df.buyer.map(buyer)
    ints = lambda v: int(float(v)) if str(v).strip().replace(".", "").isdigit() else str(v).strip()
    for c in ("order", "suffix", "line"): df[c] = df[c].map(ints)
    df["disp"] = df.disp.astype(str).str.strip(); df["vendor"] = df.vendor.astype(str).str.strip()
    df["comm"] = df.comm.map(lambda v: None if pd.isna(v) else (v if isinstance(v, (pd.Timestamp, dt.datetime)) else str(v).strip()))
    bad = df[df.task.astype(str).str.strip().str.lower() != "past due"]
    for v in sorted(set(bad.task.astype(str))): w.append(f"TASK '{v}' is not 'Past Due' - changed to 'Past Due'")
    d = pd.Timestamp(date).normalize()
    return d, df, w

def _cell(ref, v, style):
    s = f' s="{style}"' if style else ""
    if v is None or (isinstance(v, float) and pd.isna(v)) or v == "": return ""
    if isinstance(v, (dt.datetime, pd.Timestamp)): return f'<c r="{ref}"{s}><v>{(pd.Timestamp(v).to_pydatetime() - EPOCH).days}</v></c>'
    if isinstance(v, (int, float)) and not isinstance(v, bool): return f'<c r="{ref}"{s}><v>{v}</v></c>'
    return f'<c r="{ref}"{s} t="inlineStr"><is><t>{escape(str(v))}</t></is></c>'

def update_report(report_bytes, jobs):
    """jobs: list of (eod_file, agent_name, date). Returns (new_xlsx_bytes, log)."""
    ctx = learn(report_bytes); log = []; frames = []
    for f, agent, date in jobs:
        if agent not in ctx["agents"]: log.append(f"WARNING agent '{agent}' is not a column in the report's Daily sheet")
        d, df, w = clean(f, agent, date); log += [f"WARNING {agent}: {x}" for x in sorted(set(w))]
        if (int((d.to_pydatetime() - EPOCH).days), agent) in ctx["done"]: raise ValueError(f"Already loaded: {agent} for {d:%d-%b-%Y}. Nothing was added.")
        frames.append((d, agent, df))
    if len({(d, a) for d, a, _ in frames}) < len(frames): raise ValueError("The same agent and date was selected twice.")
    zin = zipfile.ZipFile(io.BytesIO(report_bytes)); paths = _paths(zin); dp = paths[DATA]; st = ctx["styles"]
    latest = max(d for d, _, _ in frames); n_new = sum(len(df) for _, _, df in frames)
    out = io.BytesIO(); zo = zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED)
    msheet = month_label(latest, ctx["sheets"]); small = {}
    if msheet in paths:
        x = zin.read(paths[msheet]).decode(); x, k = re.subn(r'(<c r="B6"[^>]*>)<v>\d+</v>', lambda m: f'{m.group(1)}<v>{(latest.to_pydatetime() - EPOCH).days}</v>', x, count=1)
        small[paths[msheet]] = x; log.append(f"{msheet}!B6 set to {latest:%d-%b-%Y}" if k else f"WARNING could not set {msheet}!B6")
    else: log.append(f"WARNING there is no '{msheet}' sheet yet - copy last month's sheet, rename it, then upload again")
    wbx = zin.read("xl/workbook.xml").decode()
    small["xl/workbook.xml"] = re.sub(r"<calcPr[^>]*/>", '<calcPr calcId="191029" fullCalcOnLoad="1"/>', wbx) if "<calcPr" in wbx else wbx.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
    for n in zin.namelist():   # make pivot tables refresh when the file is opened
        if re.match(r"xl/pivotCache/pivotCacheDefinition\d+\.xml$", n):
            x = zin.read(n).decode(); small[n] = x if "refreshOnLoad" in x else x.replace("<pivotCacheDefinition ", '<pivotCacheDefinition refreshOnLoad="1" ', 1)
    for it in zin.infolist():
        if it.filename == "xl/calcChain.xml": continue
        if it.filename != dp:
            zo.writestr(it, small.get(it.filename, None) if it.filename in small else zin.read(it.filename)); continue
        zi = zipfile.ZipInfo(it.filename, it.date_time); zi.compress_type = zipfile.ZIP_DEFLATED
        with zin.open(it) as src, zo.open(zi, "w", force_zip64=True) as dst:
            first = src.read(CH); m = re.search(rb'<dimension ref="A1:([A-Z]+)(\d+)"', first); last_row = int(m.group(2))
            first = first.replace(m.group(0), b'<dimension ref="A1:%s%d"' % (m.group(1), last_row + n_new), 1)
            xr = []; r = last_row
            for d, agent, df in sorted(frames, key=lambda t: (t[0], t[1])):
                for x in df.itertuples(index=False):
                    r += 1; vals = [month_label(d, ctx["sheets"]), d, week_of(d), agent, TASK, x.order, x.suffix, x.line, x.vendor, x.buyer, x.disp, x.comm, None, None]
                    xr.append(f'<row r="{r}">' + "".join(_cell(f"{LET[j]}{r}", v, st.get(LET[j])) for j, v in enumerate(vals)) + "</row>")
            rows_xml = "".join(xr).encode(); carry = b""; data = first
            while True:
                data = carry + data; dst.write(data[:-KEEP]); carry = data[-KEEP:]
                data = src.read(CH)
                if not data: break
            if f'<row r="{last_row}"'.encode() not in carry: raise ValueError("Report layout changed: last Data row not where expected. Nothing was added.")
            i = carry.rindex(b"</sheetData>"); dst.write(carry[:i] + rows_xml + carry[i:])
    zo.close()
    for d, a, df in sorted(frames, key=lambda t: (t[0], t[1])): log.append(f"{a}: +{len(df)} rows for {d:%d-%b-%Y}")
    log.insert(0, f"Data: +{n_new} rows (now {last_row + n_new})")
    return out.getvalue(), log

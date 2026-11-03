"""Ballast water: the per-tank table of the master sheet 'Ballast' and the tank-history table of the reporting forms."""
import re
from openpyxl.cell.cell import MergedCell
from matching import norm

KEYS = [  # (key, regex on the normalised master header)
    ("tank", r"^tank"), ("cur_vol", r"current volume"),
    ("src_date", r"last ballasting date"), ("src_port", r"last ballasting port"), ("src_sal", r"last ballasting salinity"),
    ("mg_date", r"management date"), ("mg_start", r"management start"), ("mg_end", r"management end"),
    ("mg_vol", r"management volume"), ("mg_pct", r"management .*exch"), ("mg_method", r"management method"),
    ("mg_wave", r"management wave"),
    ("ds_date", r"discharge date"), ("ds_port", r"discharge port"), ("ds_vol", r"discharge volume"),
    ("ds_sal", r"discharge salinity"),
]


def read_master(wb):
    for name in ("Ballast_Tanks", "Ballast"):
        if name in wb.sheetnames:
            rows = _read_sheet(wb[name])
            if rows:
                return rows
    return []


def _read_sheet(ws):
    hdr_row, cols = None, {}
    for r in range(1, min(ws.max_row, 80) + 1):
        t = norm(str(ws.cell(r, 2).value or ""))
        if t.startswith("tank") and not t.startswith("tanks "):
            c2 = {}
            for c in range(2, ws.max_column + 1):
                h = norm(str(ws.cell(r, c).value or ""))
                for key, rx in KEYS:
                    if key not in c2.values() and re.search(rx, h):
                        c2[c] = key
                        break
            if len(c2) >= 4:
                hdr_row, cols = r, c2
                break
    if not hdr_row:
        return []
    rows = []
    for r in range(hdr_row + 1, ws.max_row + 1):
        rec = {k: ws.cell(r, c).value for c, k in cols.items()}
        if rec.get("tank") in (None, ""):
            continue
        rows.append({k: v for k, v in rec.items() if v not in (None, "")})
    return rows


def _num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def derive_totals(master):
    """Totals that are not typed into the master are worked out from the tank table."""
    t = master.ballast_tanks
    if not t:
        return
    f = master.fields
    vols = [_num(x.get("cur_vol")) for x in t]
    if "bw_on_board" not in f:
        s = sum(v for v in vols if v)
        f["bw_on_board"] = int(s) if float(s).is_integer() else round(s, 3)
    if "bw_in_ballast" not in f:
        f["bw_in_ballast"] = sum(1 for v in vols if v)
    if "bw_ds_tanks" not in f:
        f["bw_ds_tanks"] = sum(1 for x in t if (_num(x.get("ds_vol")) or 0) > 0)
    meth = [str(x.get("mg_method") or "").upper() for x in t]
    if "bw_exchanged" not in f:
        f["bw_exchanged"] = sum(1 for m in meth if m in ("ER", "FT"))
    if "bw_alt" not in f:
        f["bw_alt"] = sum(1 for m in meth if m == "ALT")


# ------------------------------------------------------------------ the form
GROUPS = [("cur", r"current volume"), ("src", r"last ballasting|bw source|ballast source|source"),
          ("mg", r"management|bw exchange|exchange"), ("ds", r"proposed|discharge")]
SUB = [("date", r"date|tarih|datum|fecha|data\b"), ("port", r"port|lat|liman|hafen|puerto|porto"),
       ("start", r"start"), ("end", r"end point"), ("sal", r"salinity|tuzluluk|salzgehalt|salinidad|salinidade"),
       ("temp", r"temp|sicaklik"), ("pct", r"%|exch"), ("method", r"method|metod"), ("wave", r"wave"),
       ("vol", r"volume|miktar|menge|volumen")]
MAP = {("cur", None): "cur_vol", ("src", "date"): "src_date", ("src", "port"): "src_port", ("src", "sal"): "src_sal",
       ("src", "vol"): "cur_vol", ("mg", "date"): "mg_date", ("mg", "start"): "mg_start", ("mg", "end"): "mg_end",
       ("mg", "port"): "mg_end", ("mg", "vol"): "mg_vol", ("mg", "pct"): "mg_pct", ("mg", "method"): "mg_method",
       ("mg", "wave"): "mg_wave", ("mg", "sal"): "mg_sal", ("ds", "date"): "ds_date", ("ds", "port"): "ds_port",
       ("ds", "vol"): "ds_vol", ("ds", "sal"): "ds_sal"}


def find_table(ws):
    """(first data row, {column: key}) of a tank-history table, or None."""
    from forms import span
    tank = None
    for r in range(1, min(ws.max_row, 80) + 1):
        for c in range(1, min(ws.max_column, 8) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and re.match(r"(?i)^\s*(\d+\.?\s*)?tanks?\s*/\s*holds", v.strip()):
                tank = (r, c)
                break
        if tank:
            break
    if not tank:
        return None
    r0, c0 = tank
    _, _, rr0, rr1 = span(ws, r0, c0)
    # group labels: cells in the rows of the tank header to the right of it
    groups = []
    for r in range(rr0, rr1 + 1):
        for c in range(c0 + 1, min(ws.max_column, 60) + 1):
            cell = ws.cell(r, c)
            v = cell.value
            if isinstance(cell, MergedCell) or not isinstance(v, str):
                continue
            t = norm(v)
            for g, rx in GROUPS:
                if re.search(rx, t):
                    groups.append((c, span(ws, r, c)[1], g, r))
                    break
    groups.sort()
    if not groups:
        return None
    # header rows: from the group row down to the end of the tank header merge (or the first rows with sub-labels)
    hdr_last = max(rr1, groups[0][3])
    sub_rows = []
    for r in range(groups[0][3], groups[0][3] + 6):
        n = 0
        for c in range(c0 + 1, min(ws.max_column, 60) + 1):
            cell = ws.cell(r, c)
            if isinstance(cell, MergedCell) or not isinstance(cell.value, str):
                continue
            t = norm(cell.value)
            if any(re.search(rx, t) for _, rx in SUB) and not any(re.search(rx, t) for _, rx in GROUPS):
                n += 1
        if n >= 2:
            sub_rows.append(r)
    if sub_rows:
        hdr_last = max(hdr_last, max(sub_rows) + 0)
    # translation row right under the English sub-headers (text but no recognisable word)
    r = hdr_last + 1
    while r <= hdr_last + 2 and any(isinstance(ws.cell(r, c).value, str) and not isinstance(ws.cell(r, c), MergedCell)
                                    for c in range(c0 + 1, min(ws.max_column, 60) + 1)) \
            and not any(isinstance(ws.cell(r, cc).value, (int, float)) for cc in range(c0, c0 + 4)):
        hdr_last = r
        r += 1
    cols = {}
    for g_c0, g_c1, g, g_row in groups:
        pass
    for i, (g_c0, g_c1, g, g_row) in enumerate(groups):
        end = groups[i + 1][0] - 1 if i + 1 < len(groups) else min(ws.max_column, 60)
        end = max(end, g_c1)
        found = False
        for r in sub_rows:
            for c in range(g_c0, end + 1):
                cell = ws.cell(r, c)
                if isinstance(cell, MergedCell) or not isinstance(cell.value, str):
                    continue
                t = norm(cell.value)
                for s, rx in SUB:
                    if re.search(rx, t) and (g, s) in MAP:
                        cols.setdefault(c, MAP[(g, s)])
                        found = True
                        break
        if not found and (g, None) in MAP:
            cols[g_c0] = MAP[(g, None)]
    cols[c0] = "tank"
    return hdr_last + 1, cols, r0


def fill_tanks(ctx, ws, res):
    from forms import anchor, blank, is_formula
    import convert as C
    tab = find_table(ws)
    if not tab:
        return None
    first, cols, hdr0 = tab
    tanks = ctx["master"].ballast_tanks
    # rows of the table: until the legend line / a numbered line
    last = first - 1
    for r in range(first, min(ws.max_row, first + 120) + 1):
        a = ws.cell(r, min(cols)).value
        row_text = " ".join(str(ws.cell(r, c).value) for c in range(1, min(ws.max_column, 12) + 1) if ws.cell(r, c).value)
        if re.search(r"(?i)tank codes|^\s*\d+\.\s|responsible officer|signature", row_text) and r > first:
            break
        last = r
    rows = list(range(first, last + 1))
    done = 0
    for i, r in enumerate(rows):
        rec = tanks[i] if i < len(tanks) else None
        for c, key in cols.items():
            cell = anchor(ws, r, c)
            if isinstance(cell, MergedCell) or is_formula(cell.value):
                continue
            if rec is None:
                if not blank(cell.value):
                    cell.value = None
                continue
            val = rec.get(key)
            if key in ("src_date", "mg_date", "ds_date") and val not in (None, ""):
                d = C.to_date(val)
                if d:
                    if not C.is_date_format(cell.number_format):
                        cell.number_format = "dd-mm-yyyy"
                    val = d
            elif key in ("cur_vol", "mg_vol", "ds_vol", "src_sal", "ds_sal", "mg_sal", "mg_pct", "mg_wave"):
                n = _num(val)
                val = (int(n) if n is not None and float(n).is_integer() else n) if n is not None else val
            cell.value = val if val not in ("",) else None
        if rec is not None:
            done += 1
    if len(tanks) > len(rows):
        res.note(f"{len(tanks) - len(rows)} tank(s) did not fit in the table of this form: add them on a second page.")
    if not tanks:
        res.note("The master sheet 'Ballast' has no tank table: the tank history was emptied - fill it in by hand.")
    res.lines.append(f"{done} of {len(tanks)} tank lines filled")
    return rows, cols, hdr0

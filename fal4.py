"""IMO FAL 4 (crew's effects): list with surname / given names (often linked by formulas) + cigarettes, tobacco, wines."""
import re
from openpyxl.cell.cell import MergedCell
from matching import norm

COLS = [("cig", r"^cigarettes?$"), ("tob", r"^tobacco$"), ("wine", r"^wines?( spirits)?$|^wines? spirits$"),
        ("other", r"^other goods$|^other$"), ("fam", r"^(sur|family) ?name$"), ("giv", r"^given names?$|^first name"),
        ("rank", r"^rank")]


def find(ws):
    for r in range(1, min(ws.max_row, 30) + 1):
        cols = {}
        for c in range(1, min(ws.max_column, 14) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str):
                t = norm(v)
                for k, rx in COLS:
                    if re.search(rx, t):
                        cols[k] = c
                        break
        if {"cig", "tob", "wine"} <= set(cols):
            return r, cols
    return None


def fill_fal4(ctx, ws, res):
    from forms import mark, anchor, is_formula, blank
    hit = find(ws)
    if not hit:
        return False
    hr, cols = hit
    m = ctx["master"]
    f = m.fields
    n = 0
    # header fields
    port = f.get("port_locode")
    eta = f.get("eta_arrival") or f.get("eta")
    for r in range(1, hr):
        for c in range(1, min(ws.max_column, 12) + 1):
            v = ws.cell(r, c).value
            if not isinstance(v, str):
                continue
            t = norm(v)
            val = None
            if t.startswith("un locode of arrival"):
                val = port
            elif t.startswith("un locode of departure"):
                val = f.get("last_port_locode") if str(f.get("movement", "")).lower().startswith("arr") else port
            elif t == "name of ship":
                val = f.get("ship_name")
            elif t == "imo number":
                val = f.get("imo")
            elif t == "call sign":
                val = f.get("call_sign")
            elif t.startswith("flag state"):
                val = f.get("flag")
            elif t.startswith("time of arrival") and eta is not None and hasattr(eta, "hour"):
                val = eta.strftime("%H:%M:%S")
            if val is None:
                continue
            tgt = anchor(ws, r, c + 1)
            if tgt is not None and not is_formula(tgt.value) and not isinstance(tgt, MergedCell):
                tgt.value = val
                n += 1
    r0 = hr + 1
    tracked = {"cig": "eff_cig", "tob": "eff_tob", "wine": "eff_wine"}
    for i, rec in enumerate(m.crew):
        row = r0 + i
        for k, key in tracked.items():
            cell = ws.cell(row, cols[k])
            if isinstance(cell, MergedCell) or is_formula(cell.value):
                continue
            val = rec.get(key)
            if blank(val):
                if not blank(cell.value):
                    cell.value = None
                continue
            cell.value = val
            n += 1
        if "other" in cols:
            cell = ws.cell(row, cols["other"])
            other = " ".join(str(x) for x in (rec.get("eff_cash"), rec.get("eff_elec")) if x)
            if other and not isinstance(cell, MergedCell) and not is_formula(cell.value):
                cell.value = other
                n += 1
        for k in ("fam", "giv", "rank"):
            if k in cols:
                cell = ws.cell(row, cols[k])
                if not isinstance(cell, MergedCell) and not is_formula(cell.value) and blank(cell.value):
                    cell.value = rec.get({"fam": "family_name", "giv": "given_names", "rank": "rank"}[k])
    res.kind = "FAL 4 Crew's effects"
    res.lines.append(f"{len(m.crew)} crew lines checked, {n} values filled")
    return True

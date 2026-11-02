"""IMO FAL 4 (crew's effects): list with surname / given names (often linked by formulas) + cigarettes, tobacco, wines."""
import re
from openpyxl.cell.cell import MergedCell
from matching import norm

COLS = [("cig", r"^cigarettes?( pcs)?$"), ("tob", r"^tobacco( kg)?$"), ("wine", r"^wines?( spirits)?$|^wines? spirits$|^spirits( l)?$"),
        ("other", r"^other goods$|^other$"), ("fam", r"^(sur|family) ?name$"), ("giv", r"^given names?$|^first name"),
        ("full", r"^family name given names?$|^name$"), ("rank", r"^rank"), ("no", r"^(list )?(no|number)$")]


def find(ws):
    for r in range(1, min(ws.max_row, 30) + 1):
        cols = {}
        for c in range(1, min(ws.max_column, 14) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str):
                t = re.sub(r"^\d+ ", "", norm(v))
                for k, rx in COLS:
                    if re.search(rx, t):
                        cols[k] = c
                        break
        if {"cig", "tob", "wine"} <= set(cols):
            return r, cols
        if "cig" in cols and "wine" in cols:        # German customs form: header spread over several lines
            for r2 in range(r, r + 4):
                for c in range(1, min(ws.max_column, 14) + 1):
                    v = ws.cell(r2, c).value
                    if isinstance(v, str):
                        t = re.sub(r"^\d+ ", "", norm(v))
                        for k, rx in COLS:
                            if k in ("full", "rank", "no", "fam", "giv") and re.search(rx, t) and k not in cols:
                                cols[k] = c
            last = max(r2 for r2 in range(r, r + 4) if any(isinstance(ws.cell(r2, c).value, str) and re.search(r"family name|rank|^\d+\. no", ws.cell(r2, c).value, re.I) for c in range(1, 14))) \
                if any(isinstance(ws.cell(r2, c).value, str) and re.search(r"family name|rank", ws.cell(r2, c).value, re.I) for r2 in range(r, r + 4) for c in range(1, 14)) else r
            cols["_lenient"] = True
            return last, cols
    return None


def fill_fal4(ctx, ws, res):
    from forms import mark, anchor, is_formula, blank
    hit = find(ws)
    if not hit:
        return False
    hr, cols = hit
    lenient = cols.pop("_lenient", False)
    if lenient:
        from stores import fill_header
        fill_header(ctx, ws, res, hr)
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
    tracked = {k: v for k, v in tracked.items() if k in cols}
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
        if "full" in cols:
            cell = ws.cell(row, cols["full"])
            if not isinstance(cell, MergedCell) and not is_formula(cell.value) and blank(cell.value):
                cell.value = ", ".join(str(x) for x in (rec.get("family_name"), rec.get("given_names")) if x) or None
        if "no" in cols:
            cell = ws.cell(row, cols["no"])
            if not isinstance(cell, MergedCell) and not is_formula(cell.value) and blank(cell.value):
                cell.value = i + 1
        for k in ("fam", "giv", "rank"):
            if k in cols:
                cell = ws.cell(row, cols[k])
                if not isinstance(cell, MergedCell) and not is_formula(cell.value) and blank(cell.value):
                    cell.value = rec.get({"fam": "family_name", "giv": "given_names", "rank": "rank"}[k])
    res.kind = "FAL 4 Crew's effects"
    res.lines.append(f"{len(m.crew)} crew lines checked, {n} values filled")
    return True

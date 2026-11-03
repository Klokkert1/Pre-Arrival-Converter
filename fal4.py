"""IMO FAL 4 (crew's effects): list with surname / given names (often linked by formulas) + cigarettes, tobacco, wines."""
import re
from openpyxl.cell.cell import MergedCell
from matching import norm

COLS = [("cig", r"^cigarettes?( (pcs|pieces|stk|st|number|nr|qty|stuks|sticks))?$"),
        ("tob", r"^t[oa]b[a]?cco( (kg|gr|g|grams?|gram|kgs))?$"),
        ("spirits", r"^spirits?( (l|ltr|litres?|liters?))?$"),
        ("wine", r"^wines?( (spirits|l|ltr|litres?|liters?))?$|^wines? spirits$"),
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
        if "spirits" in cols and "wine" not in cols:
            cols["wine"] = cols.pop("spirits")
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
    import convert as C
    from forms import mark, anchor, is_formula, blank
    hit = find(ws)
    if not hit:
        return False
    hr, cols = hit
    lenient = cols.pop("_lenient", False)
    numbered = any(isinstance(ws.cell(r, c).value, str)
                   and re.match(r"\s*\d+(\.\d+)?\.?\s*(name of ship|imo n|port of arrival|date of arrival|flag state)",
                                ws.cell(r, c).value, re.I)
                   for r in range(1, hr) for c in range(1, 14))
    if lenient or numbered:
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
            elif t.startswith("date of arrival") and eta is not None:
                val = C.to_date(eta)
            elif t.startswith("date of departure") and f.get("etd_departure") is not None:
                val = C.to_date(f.get("etd_departure"))
            elif t.startswith("time of departure") and hasattr(f.get("etd_departure"), "hour"):
                val = f.get("etd_departure").strftime("%H:%M:%S")
            if val is None:
                continue
            tgt = anchor(ws, r, c + 1)
            if tgt is not None and not is_formula(tgt.value) and not isinstance(tgt, MergedCell):
                tgt.value = val
                if hasattr(val, "year") and not C.is_date_format(tgt.number_format):
                    tgt.number_format = "dd-mm-yyyy"
                n += 1
    r0 = hr + 1
    # how many person lines does this sheet have, and which crew members belong on it (multi-page forms)?
    cap = 0
    for rr in range(r0, min(ws.max_row, r0 + 80) + 1):
        txt = " ".join(str(ws.cell(rr, cc).value) for cc in range(1, 14) if isinstance(ws.cell(rr, cc).value, str))
        if re.search(r"date and signature|imo fal|^\s*\d+\.\s*date|signature of master", txt, re.I):
            break
        cap += 1
    if cap == 0:
        cap = len(m.crew)
    paged = re.search(r"page\s*\d+", ws.title, re.I)
    off = 0
    if paged:
        off = ctx.get("fal4_paged_used", 0)
        ctx["fal4_paged_used"] = off + cap
    crew_here = m.crew[off:off + cap]
    tracked = {"cig": "eff_cig", "tob": "eff_tob", "wine": "eff_wine"}
    tracked = {k: v for k, v in tracked.items() if k in cols}
    if "spirits" in cols:
        res.note("This form has separate Spirits and Wines columns; the master has one combined field, "
                 "which is placed under Wines. Check the Spirits column by hand.")
    for j in range(len(crew_here), cap):          # lines without a crew member: remove old content
        for k in ("fam", "giv", "rank", "full", "no", "cig", "tob", "wine", "spirits", "other"):
            if k in cols:
                cell = ws.cell(r0 + j, cols[k])
                if not isinstance(cell, MergedCell) and not is_formula(cell.value) and not blank(cell.value):
                    cell.value = None
    for i, rec in enumerate(crew_here):
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
            if not isinstance(cell, MergedCell) and not is_formula(cell.value):
                cell.value = ", ".join(str(x) for x in (rec.get("family_name"), rec.get("given_names")) if x) or None
        if "no" in cols:
            cell = ws.cell(row, cols["no"])
            if not isinstance(cell, MergedCell) and not is_formula(cell.value):
                cell.value = i + 1 + off
        for k in ("fam", "giv", "rank"):
            if k in cols:
                cell = ws.cell(row, cols[k])
                if not isinstance(cell, MergedCell) and not is_formula(cell.value):
                    cell.value = rec.get({"fam": "family_name", "giv": "given_names", "rank": "rank"}[k])
    res.kind = "FAL 4 Crew's effects"
    res.lines.append(f"{len(crew_here)} crew lines checked, {n} values filled")
    return True


# ------------------------------------------------------------------ crew temperature list (names in pairs of AM/PM lines)
def is_templist(ws):
    for r in range(1, 9):
        for c in range(1, 16):
            v = ws.cell(r, c).value
            if isinstance(v, str) and re.search(r"temperature list|temp\.? list", v, re.I):
                return True
    return False


def fill_templist(ctx, ws, res):
    from forms import anchor, is_formula
    from stores import fill_header
    m = ctx["master"]
    hdr = None
    for r in range(5, 20):
        for c in range(1, 12):
            v = ws.cell(r, c).value
            if isinstance(v, str) and re.search(r"family name", v, re.I):
                hdr = (r, c)
                break
        if hdr:
            break
    if not hdr:
        return False
    hr, hc = hdr
    n = fill_header(ctx, ws, res, hr)
    # person lines: two lines per person (AM / PM), up to the signature line
    rows = []
    r = hr + 2
    while r <= ws.max_row:
        txt = " ".join(str(ws.cell(r, cc).value) for cc in range(1, 16) if isinstance(ws.cell(r, cc).value, str))
        if re.search(r"date and signature|signature", txt, re.I):
            break
        rows.append(r)
        r += 2
    cap = len(rows)
    paged = re.search(r"page\s*\d+", ws.title, re.I)
    off = 0
    if paged:
        off = ctx.get("temp_paged_used", 0)
        ctx["temp_paged_used"] = off + cap
    crew_here = m.crew[off:off + cap]
    for j, row in enumerate(rows):
        no_c = anchor(ws, row, hc)
        name_c = anchor(ws, row, hc + 1)
        if j < len(crew_here):
            rec = crew_here[j]
            if not is_formula(no_c.value) and not isinstance(no_c, MergedCell):
                no_c.value = off + j + 1
            if not is_formula(name_c.value) and not isinstance(name_c, MergedCell):
                name_c.value = ", ".join(str(x) for x in (rec.get("family_name"), rec.get("given_names")) if x)
            n += 1
        else:
            for cell in (no_c, name_c):
                if not isinstance(cell, MergedCell) and not is_formula(cell.value):
                    cell.value = None
    res.kind = "Crew temperature list"
    res.lines.append(f"{n} fields and {len(crew_here)} crew names filled; temperatures are not in the master and stay empty")
    return True

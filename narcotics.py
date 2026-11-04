"""Narcotics (medical chest) list: one line per item from the master sheet 'Narcotics'."""
import re
from openpyxl.cell.cell import MergedCell
from matching import norm


def read_master(wb):
    """Rows of the master sheet 'Narcotics' (header row 5, data from row 6; example lines are skipped)."""
    if "Narcotics" not in wb.sheetnames:
        return []
    ws = wb["Narcotics"]
    hdr = {}
    hr = 5
    for r in range(1, 31):       # header row: wherever the ship header put it
        if any(norm(str(ws.cell(r, c).value or "")).startswith("name of narcotic") for c in range(1, ws.max_column + 1)):
            hr = r
            break
    for c in range(1, ws.max_column + 1):
        h = norm(str(ws.cell(hr, c).value or ""))
        if not h:
            continue
        for key, rx in (("name", r"^name"), ("strength", r"strength|form"), ("qty", r"quantity|stock|on board"),
                        ("unit", r"^unit"), ("expiry", r"expir"), ("place", r"stored|location|place"),
                        ("remarks", r"remark"), ("code", r"article code|artikel code|artical code"),
                        ("number", r"article number|artikel number|artical number|article no")):
            if re.search(rx, h) and key not in hdr:
                hdr[key] = c
    rows = []
    for r in range(hr + 1, ws.max_row + 1):
        rec = {k: ws.cell(r, c).value for k, c in hdr.items()}
        name = rec.get("name")
        if name in (None, "") or str(name).strip().lower() in ("extra",) or str(name).upper().startswith("EXAMPLE"):
            continue
        rows.append({k: v for k, v in rec.items() if v not in (None, "")})
    return rows


def _header(ws):
    for r in range(1, min(ws.max_row, 40) + 1):
        cols = {}
        for c in range(1, min(ws.max_column, 14) + 1):
            v = ws.cell(r, c).value
            if not isinstance(v, str):
                continue
            t = norm(v)
            pass
        texts = {c: norm(str(ws.cell(r, c).value)) for c in range(1, min(ws.max_column, 14) + 1)
                 if isinstance(ws.cell(r, c).value, str)}
        if any("stock" in t or "quantity" in t for t in texts.values()) and any(
                re.search(r"name|description", t) for t in texts.values()) and any(
                re.match(r"^(\d+ )?(no|nr|number)$", t) or t.startswith("no ") for t in texts.values()):
            for c, t in texts.items():
                if re.search(r"article code|artical code|artikel code|^\d* ?code", t):
                    cols["code"] = c
                elif re.search(r"article number|artical number|artikel number|artical no|article no", t):
                    cols["number"] = c
                elif re.search(r"stock|quantity|on board", t):
                    cols.setdefault("qty", c)
                elif re.search(r"expir", t):
                    cols["expiry"] = c
                elif re.search(r"stored|location", t):
                    cols["place"] = c
                elif re.search(r"remark", t):
                    cols["remarks"] = c
                elif re.search(r"strength", t):
                    cols["strength"] = c
                elif re.search(r"^unit", t):
                    cols["unit"] = c
                elif re.search(r"name|description|drug|narcotic", t) and "name" not in cols:
                    cols["name"] = c
                elif re.match(r"^(\d+ )?(no|nr|number)$", t) or t.startswith("no "):
                    cols.setdefault("serial", c)
            if "name" in cols and "qty" in cols:
                return r, cols
    return None, None


def is_narcotic(ws):
    text = (ws.title + " " + " ".join(str(ws.cell(r, c).value) for r in range(1, 6) for c in range(1, 5)
                                      if ws.cell(r, c).value)).lower()
    return "narcotic" in text and _header(ws)[0] is not None


def fill_narcotics(ctx, ws, res):
    from forms import mark, anchor, blank, is_formula
    r0, cols = _header(ws)
    items = ctx["master"].narcotics
    res.kind = "Narcotics list"
    # data lines: from the header down to the end of the numbered / bordered lines
    last = r0
    for r in range(r0 + 1, min(ws.max_row, r0 + 200) + 1):
        a = ws.cell(r, cols.get("serial", 1)).value
        if a is None and all(blank(ws.cell(r, c).value) for c in cols.values()):
            if r > last + 3:
                break
            continue
        if isinstance(a, str) and re.search(r"signature|master|date|place", a, re.I) and not re.match(r"^\d+$", a.strip()):
            break
        last = r
    rows = list(range(r0 + 1, last + 1))
    n = 0
    for i, r in enumerate(rows):
        item = items[i] if i < len(items) else None
        for key in ("name", "code", "number", "qty", "expiry", "place", "remarks", "strength", "unit"):
            if key not in cols:
                continue
            cell = anchor(ws, r, cols[key])
            if isinstance(cell, MergedCell) or is_formula(cell.value):
                continue
            if item is None:
                if not blank(cell.value):
                    cell.value = None
                continue
            if key == "name":
                val = " ".join(str(x) for x in (item.get("name"), item.get("strength")) if x) if "strength" not in cols else item.get("name")
            elif key == "qty":
                q = item.get("qty")
                if q is None:
                    val = None
                else:
                    try:
                        q = int(q) if float(q).is_integer() else q
                    except (TypeError, ValueError):
                        pass
                    val = q if "unit" in cols else f"{q} {item.get('unit') or ''}".strip()
            elif key == "expiry":
                import convert as C
                d = C.to_date(item.get("expiry"))
                val = d if d else item.get("expiry")
                if d and not C.is_date_format(cell.number_format):
                    cell.number_format = "dd-mm-yyyy"
            else:
                val = item.get(key)
            if val is None and key in ("code", "number"):
                mark(cell)
                res.yellow += 1
                if key == "code" and "Article code" not in res.missing:
                    res.missing.append("Article code")
                if key == "number" and "Article number" not in res.missing:
                    res.missing.append("Article number")
                cell.value = None
                continue
            cell.value = val
        if item is not None:
            n += 1
    if len(items) > len(rows):
        res.note(f"{len(items) - len(rows)} narcotic item(s) did not fit on this page; use the next page or add lines.")
    if not items:
        res.note("The master sheet 'Narcotics' has no items: the list was emptied - fill in by hand if needed.")
    res.lines.append(f"{n} of {len(items)} narcotic lines filled")
    return rows, cols

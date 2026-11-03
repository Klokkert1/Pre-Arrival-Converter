"""Ship's stores (FAL 3) with a printed article list: header fields and one line per article."""
import re
from openpyxl.cell.cell import MergedCell
from matching import norm, compact

ART = r"^(\d+\.?\d*\.?\s*)?(name of article|number of article|article|articles|name of the article|artikel|artikelbezeichnung|nombre del articulo|description of (the )?stores)\b"
QTY = ("quantity", "qty", "menge", "cantidad", "quantity kg ltr")
UNIT = ("unit", "units", "einheit", "unidad")
LOC = r"location on board|place of storage|location|lagerort|ubicacion|stowage"
STOP = re.compile(r"signature|date and sig|^\s*\d+\.\s*date", re.I)
NUMLABEL = re.compile(r"^\s*\d{1,2}(\.\d{1,2})?\.?\s+\S")

HEADER = [
    ("loc_arr", r"^un ?locode of arrival"),
    ("loc_dep", r"^un ?locode of departure"),
    ("date_arr", r"^date of arrival$"),
    ("date_dep", r"^date of departure$"),
    ("time_arr", r"^time of arrival"),
    ("time_dep", r"^time of departure"),
    ("last_poc", r"^last port of call"),
    ("next_poc", r"^next port of call"),
    ("stay", r"^period of stay"),
    ("ship_name", r"^name of ship|^name des schiffes|^nombre del buque"),
    ("imo", r"^imo number|^imo no|^imo nummer"),
    ("call_sign", r"^call sign|^rufzeichen|^indicativo"),
    ("voyage_no", r"^voyage number|^reisenummer"),
    ("port", r"^port of arrival ?/ ?departure|^port of (arrival|departure)"),
    ("date", r"^date of arrival ?/ ?departure|^date of (arrival|departure)"),
    ("flag", r"^flag state|^flag of ship|^flagge|^nationality of ship"),
    ("from_dest", r"^port arrived from|port of destination"),
    ("lastnext", r"last port of call ?/ ?next port of call|last/next port"),
    ("persons", r"^number of persons on board|^persons on board"),
]


def _blank(v):
    return v is None or (isinstance(v, str) and not v.strip())


def find_article_table(ws):
    """Returns dict(row, art, qty, unit, loc) or None."""
    for r in range(1, min(ws.max_row, 40) + 1):
        art = None
        for c in range(1, min(ws.max_column, 12) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and re.search(ART, re.sub(r"\s+", " ", v.lstrip("* ")).strip().lower()):
                art = c
                break
        if not art:
            continue
        cols = {"row": r, "art": art, "qty": None, "unit": None, "loc": None}
        for c in range(art + 1, min(ws.max_column, art + 10) + 1):
            v = ws.cell(r, c).value
            if not isinstance(v, str):
                continue
            t = re.sub(r"^\s*\d+\.?\d*\.?\s*", "", re.sub(r"\s+", " ", v.lstrip("* ")).strip().lower())
            if cols["qty"] is None and any(t.startswith(q) for q in QTY):
                cols["qty"] = c
            elif cols["unit"] is None and t in UNIT:
                cols["unit"] = c
            elif cols["loc"] is None and re.search(LOC, t):
                cols["loc"] = c
        if cols["qty"]:
            return cols
    return None


def _anchor(ws, r, c):
    from forms import anchor
    return anchor(ws, r, c)


def _num(x):
    return int(x) if float(x).is_integer() else round(float(x), 2)


def _aggregate(items, unit_text):
    """(value, unit, ok): sum of the master items expressed in unit_text (or in the first item's unit)."""
    from forms import UNIT_FAMILY
    target = unit_text or (items[0]["unit"] if items else "")
    fam = UNIT_FAMILY.get(compact(target or ""))
    total, ok = 0.0, True
    for it in items:
        try:
            q = float(it["qty"])
        except (TypeError, ValueError):
            ok = False
            continue
        f = UNIT_FAMILY.get(compact(it["unit"] or ""))
        if fam and f and f[0] == fam[0]:
            total += q * f[1] / fam[1]
        elif norm(it["unit"] or "") == norm(target or ""):
            total += q
        else:
            ok = False
    return total, target, ok


def fill_article_table(ctx, ws, res, cols):
    from forms import store_class, mark, get_allowed
    master = ctx["master"]
    classes, names = {}, []
    for it in master.stores:
        cls = store_class(it["name"]) or store_class(it["category"] or "")
        if not cls and "alcohol" in norm(it["category"] or ""):
            cls = "other_alcohol"
        it["_cls"] = cls
        classes.setdefault(cls, []).append(it)
    used = set()
    r0 = cols["row"]
    rows = []
    r = r0 + 1
    while r <= ws.max_row:
        row_text = " ".join(str(ws.cell(r, c).value) for c in range(1, min(ws.max_column, 12) + 1)
                            if isinstance(ws.cell(r, c).value, str))
        if STOP.search(row_text) or NUMLABEL.match(str(ws.cell(r, cols["art"]).value or "")):
            break
        rows.append(r)
        r += 1
    filled = missing = 0
    blank_rows = []
    for r in rows:
        a = _anchor(ws, r, cols["art"])
        label = a.value.strip() if isinstance(a.value, str) else ""
        if not label:
            blank_rows.append(r)
            continue
        cls = store_class(re.sub(r"^\s*\d{3}\s*-\s*", "", label))
        items = [i for i in classes.get(cls, []) if id(i) not in used] if cls else []
        if not items:      # fall back to a shared word (e.g. 'Paint' in 'Paint & Thinner')
            words = {w for w in norm(label).split() if len(w) >= 4}
            items = [i for i in master.stores if id(i) not in used and words & set(norm(i["name"]).split())]
        q = _anchor(ws, r, cols["qty"])
        u = _anchor(ws, r, cols["unit"]) if cols["unit"] else None
        loc = _anchor(ws, r, cols["loc"]) if cols["loc"] else None
        if not items:
            for cell in (q, u, loc):
                if cell is not None and not _blank(cell.value):
                    cell.value = None
            mark(q); res.yellow += 1; missing += 1
            res.missing.append(label)
            continue
        for i in items:
            used.add(id(i))
        unit_text = u.value if u is not None and not _blank(u.value) else None
        if unit_text is None:
            _h = re.search(r"\((l|ltr|litres?|liters?|kg|kgs|pcs|pieces|nr|no|number)\)\s*$", label, re.I)
            if _h:
                unit_text = {"l": "L", "ltr": "L", "litre": "L", "litres": "L", "liter": "L", "liters": "L",
                             "kg": "kg", "kgs": "kg"}.get(_h.group(1).lower(), "pcs")
        total, unit, ok = _aggregate(items, unit_text)
        if not ok:
            q.value = None
            mark(q); res.yellow += 1
            res.note(f"'{label}': the unit in the master ({', '.join(sorted({str(i['unit']) for i in items}))}) "
                     f"cannot be converted safely; left empty, please fill in by hand.")
            continue
        if u is not None:
            if _blank(u.value):
                u.value = unit
                _al = get_allowed(ctx["wb"], ws, u)
                if _al:
                    import convert as _C
                    _pk = _C.choose_from_list(unit, _al)
                    if _pk is not None:
                        u.value = _pk
            q.value = _num(total)
        else:
            q.value = f"{_num(total)} {unit}".strip()
        if loc is not None:
            places = []
            for i in items:
                if i["place"] and str(i["place"]) not in places:
                    places.append(str(i["place"]))
            loc.value = ", ".join(places) if places else None
        filled += 1
    # master items that have no line on the form: use the empty lines
    rest = [i for i in master.stores if id(i) not in used]
    for it in rest:
        if not blank_rows:
            res.note(f"'{it['name']}' has no line on this form and no empty line is left - please add it by hand.")
            continue
        r = blank_rows.pop(0)
        _anchor(ws, r, cols["art"]).value = it["name"]
        try:
            qv = _num(float(it["qty"]))
        except (TypeError, ValueError):
            qv = it["qty"]
        q = _anchor(ws, r, cols["qty"])
        if cols["unit"]:
            q.value = qv
            _anchor(ws, r, cols["unit"]).value = it["unit"] or None
        else:
            q.value = f"{qv} {it['unit'] or ''}".strip()
        if cols["loc"]:
            _anchor(ws, r, cols["loc"]).value = it["place"] or None
        filled += 1
    return filled, missing, rows


# ------------------------------------------------------------------ header (numbered FAL labels, value below)
def fill_header(ctx, ws, res, upto_row):
    from forms import mark, has_border, anchor, span
    master = ctx["master"]
    f = master.fields
    movement = str(f.get("movement") or "Arrival").strip().lower()
    dep = movement.startswith("dep")
    date = f.get("etd_departure") if dep else (f.get("eta_arrival") or f.get("eta"))
    if date is None:
        date = f.get("eta")
    vals = {"ship_name": f.get("ship_name"), "imo": f.get("imo"), "call_sign": f.get("call_sign"),
            "voyage_no": f.get("voyage_no"), "port": f.get("port"), "date": date, "flag": f.get("flag"),
            "lastnext": (f"{f.get('last_port')} / {f.get('next_port')}" if f.get("last_port") or f.get("next_port") else None),
            "from_dest": (f.get("next_port") if dep else f.get("last_port")),
            "loc_arr": f.get("port_locode"),
            "loc_dep": (f.get("last_port_locode") if not dep else f.get("port_locode")),
            "date_arr": f.get("eta_arrival") or f.get("eta"), "date_dep": f.get("etd_departure"),
            "time_arr": f.get("eta_arrival") or f.get("eta"), "time_dep": f.get("etd_departure"),
            "stay": f.get("period_of_stay"),
            "persons": (len(master.crew) + len(master.pax)) if master.crew else None}
    import convert as C
    count = 0
    for r in range(1, upto_row):
        for c in range(1, min(ws.max_column, 20) + 1):
            cell = ws.cell(r, c)
            v = cell.value
            if isinstance(cell, MergedCell) or not isinstance(v, str):
                continue
            lab = re.sub(r"\s+", " ", v).strip().lower()
            lab = re.sub(r"^\d+\.?\d*\.?\s*", "", lab)
            key = next((k for k, rx in HEADER if re.search(rx, lab)), None)
            if not key:
                continue
            c0, c1, r0, r1 = span(ws, r, c)
            below = anchor(ws, r1 + 1, c0)
            right = anchor(ws, r0, c1 + 1) if c1 + 1 <= ws.max_column else None
            def labelish(x):
                if x is None or not isinstance(x.value, str):
                    return False
                if NUMLABEL.match(x.value):
                    return True
                if x.value.startswith("="):
                    return False
                t = re.sub(r"\s+", " ", x.value).strip().lower()
                t = re.sub(r"^\d+\.?\d*\.?\s*", "", t)
                if x.row >= upto_row:
                    return True
                if any(re.search(rx, t) for _k, rx in HEADER) or t.endswith(":"):
                    return True
                if re.search(r"\b(name of|date of|time of|locode|number of|period of|port of|call sign|flag state|voyage|imo)\b", t):
                    return True
                return x.value.strip().isupper() and len(x.value.split()) >= 2
            tgt = None
            for cand in (below, right):
                if cand is not None and not labelish(cand) and (not _blank(cand.value) or has_border(cand)):
                    tgt = cand
                    break
            if tgt is None:
                for cand in (right, below):
                    if cand is not None and not labelish(cand) and cand.row < upto_row:
                        tgt = cand
                        break
            if tgt is None:
                continue
            val = vals.get(key)
            if key in ("last_poc", "next_poc"):
                loc = "locode" in lab
                val = f.get(key[:4] + "_port" + ("_locode" if loc else ""))
            if key in ("time_arr", "time_dep"):
                val = val.strftime("%H:%M:%S") if hasattr(val, "hour") else None
            if _blank(val):
                mark(tgt); res.yellow += 1; res.missing.append(v.strip()[:40])
                continue
            if key in ("date", "date_arr", "date_dep"):
                d = C.to_date(val)
                if d:
                    if not C.is_date_format(tgt.number_format):
                        tgt.number_format = "dd-mm-yyyy"
                    tgt.value = d
                    count += 1
                    continue
            tgt.value = val
            count += 1
    # Arrival / Departure tick ('X' in the cell left of the chosen word)
    pos = {}
    for r in range(1, upto_row):
        for c in range(1, min(ws.max_column, 12) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip().lower() in ("arrival", "departure") and not isinstance(ws.cell(r, c), MergedCell):
                pos[v.strip().lower()] = (r, c)
    if len(pos) == 2 and pos["arrival"][0] == pos["departure"][0]:
        for word, (r, c) in pos.items():
            left = anchor(ws, r, c - 1) if c > 1 else None
            if left is None or (isinstance(left.value, str) and len(left.value.strip()) > 2):
                continue
            left.value = "X" if word == ("departure" if dep else "arrival") else None
        count += 1
    return count


def fill_stores_sheet(ctx, ws, res, cols):
    res.kind = "Ship's stores"
    n_head = fill_header(ctx, ws, res, cols["row"])
    n, missing, rows = fill_article_table(ctx, ws, res, cols)
    if not cols["unit"]:
        res.note("This form has no unit column: quantities are written with their unit (e.g. 158.4 L).")
    total = len([r for r in rows if isinstance(_anchor(ws, r, cols["art"]).value, str) and _anchor(ws, r, cols["art"]).value.strip()])
    res.lines.append(f"{n_head} header fields and {n} of {total} article lines filled"
                     + (f", {missing} not in the master" if missing else ""))

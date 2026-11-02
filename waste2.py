"""Waste declarations with bilingual / numbered layouts (Spain, Italy, ...): header fields + waste table."""
import re
from openpyxl.cell.cell import MergedCell
from matching import norm
import convert as C

IS_WASTE = re.compile(r"waste|residu|desech|garbage|abfall|basura|rifiuti", re.I)
NUMLABEL = re.compile(r"^\s*\d{1,2}(\.\d{1,2})*\.?\s+\S")
STOP = re.compile(r"confirm|signature|firma|fecha \(date\)|date and sig|i hereby|^\s*fecha\b", re.I)

COLS = [  # order matters: specific before general
    ("w_skip", r"date of delivery|time of unloading|time of delivery"),
    ("w_next_port", r"puerto en el que|port at which|next port of delivery|next delivery port|port of delivery|where (the )?remaining"),
    ("w_delivered_last", r"entregad[oa]s? en el [uú]ltimo|delivered (at|in) (the )?last|[uú]ltimo puerto de descarga|last port of (delivery|discharge)"),
    ("w_generated", r"estimaci|cantidad estimada|estimate"),
    ("w_capacity", r"capacidad|capacity|max(imum)? (dedicated )?storage"),
    ("w_onboard", r"quedan a bordo|remain|retained|on board|a bordo"),
    ("w_deliver", r"para entregar|to be delivered|to deliver|a entregarse|entregar|for delivery"),
]
LABELCOL = re.compile(r"descripci[oó]n|description")
TYPECOL = re.compile(r"^\W*(waste )?(tipo|type)\b")

CODE = [  # (regex on the lower-cased row label, master waste code)
    (r"bilge|sentina|bilgewasser", "101"),
    (r"scale and sludge|dep[oó]sitos y fangos|limpieza de (los )?tanques", "105"),
    (r"sludge|fangos|oily residues|residuos oleosos|oil residues", "102"),
    (r"tank wash|lavazas|slops|lavado de (los )?tanques", "103"),
    (r"ballast|lastre", "104"),
    (r"sewage|aguas sucias|abwasser|toilet", "401"),
    (r"plastic", "501"),
    (r"food|comida|alimento|organic", "502"),
    (r"domestic|dom[eé]stic|household", "503"),
    (r"cooking oil|aceite de cocina", "504"),
    (r"incinerator|cenizas", "505"),
    (r"operational|operacional|explotaci", "506"),
    (r"carcass|cad[aá]ver", "507"),
    (r"fishing gear|artes de pesca", "508"),
    (r"e-?waste|electr[oó]nic", "509"),
    (r"cargo residue.*(harmful|hme)|residuos de carga.*perjudic|harmful to the marine|perjudiciales", "511"),
    (r"cargo residue|restos de carga|residuos de carga", "510"),
    (r"scrap", "scrap"),
]
NLS = re.compile(r"anexo ii\b|annex ii\b|\bnls\b|noxious|nocivas", re.I)

FIELDS = [
    ("w_last_date", r"fecha de (la )?[uú]ltima entrega|date of (the )?last delivery|last delivery|last disposal date"),
    ("w_last_port", r"[uú]ltimo puerto en el que|last port (at which|where)|port of last delivery|last port waste delivered|last disposal port"),
    ("w_next_port", r"siguiente puerto de entrega|next port of delivery"),
    ("terminal", r"terminal|localidad"),
    ("eta_dt", r"estimated (date and time of )?arrival|fecha y hora estimada de (la )?llegada|\beta\b|estimated time of arrival"),
    ("etd_dt", r"estimated (date and time of )?departure|fecha y hora estimada de (la )?(partida|salida)|\betd\b|estimated time of departure"),
    ("last_port", r"[uú]ltimo puerto y pa[ií]s|last port and country"),
    ("next_port", r"siguiente puerto o pa[ií]s|next port or country"),
    ("submitter", r"persona que presenta|person submitting"),
    ("ship_name", r"nombre del buque|name of ship|^nombre\b|nome della nave"),
    ("imo", r"n[uú]mero omi|imo number|numero imo|imo no"),
    ("gt", r"arqueo bruto|gross tonnage"),
    ("ship_type", r"tipo de buque|type of ship"),
    ("owner", r"propietario|armador|owners? or operator|compa[ñn][ií]a|\bcompany\b"),
    ("call_sign", r"distintivo|call sign|indicativo"),
    ("flag", r"estado de abanderamiento|bandera|flag"),
    ("nt", r"arqueo neto|net tonnage"),
]


def blank(v):
    return v is None or (isinstance(v, str) and not v.strip())


def is_waste_sheet(ws):
    text = ws.title + " " + " ".join(str(ws.cell(r, c).value) for r in range(1, 8)
                                     for c in range(1, 8) if isinstance(ws.cell(r, c).value, str))
    return bool(IS_WASTE.search(text))


def find_table(ws):
    for r in range(1, min(ws.max_row, 14) + 1):
        cols, labelcol, typecol = {}, None, None
        for c in range(1, min(ws.max_column, 40) + 1):
            cell = ws.cell(r, c)
            v = cell.value
            if isinstance(cell, MergedCell) or not isinstance(v, str):
                continue
            t = re.sub(r"\s+", " ", v.replace("_", " ")).strip().lower()
            for k, rx in COLS:
                if re.search(rx, t):
                    cols.setdefault(c, k)
                    break
            else:
                if LABELCOL.search(t) and labelcol is None:
                    labelcol = c
                elif TYPECOL.search(t) and typecol is None:
                    typecol = c
        if len(set(cols.values())) >= 3:
            nxt = sum(1 for c in cols if isinstance(ws.cell(r + 1, c).value, str)
                      and any(re.search(rx, ws.cell(r + 1, c).value.replace("_", " ").lower()) for _, rx in COLS))
            if nxt >= 2:           # technical header row followed by a readable one (MSW templates)
                r += 1
            return {"row": r, "cols": cols, "label": labelcol or typecol or 1, "alt": typecol or 1}
    return None


def waste_code(label, section_nls=False):
    t = re.sub(r"\s+", " ", label).strip().lower()
    if NLS.search(t) or section_nls:
        return None
    m = re.search(r"\b(10[1-9]|40[1-9]|50[1-9]|51[01])\b", t)
    if m:
        return m.group(1)
    for rx, code in CODE:
        if re.search(rx, t):
            return code
    return None


def fill_table(ctx, ws, res, tb):
    from forms import anchor, span, mark, is_formula
    master = ctx["master"]
    mrows = {}
    for m in master.waste["rows"]:
        key = m["code"] if m["code"] else (waste_code(m["name"]) or None)
        if key and key not in mrows:
            mrows[key] = m
    nextport = master.waste["header"].get("w_next_port")
    vcols = tb["cols"]
    filled = cleared = unmatched = 0
    used = set()
    section_nls = False
    for r in range(tb["row"] + 1, ws.max_row + 1):
        rowtext = " ".join(str(ws.cell(r, c).value) for c in range(1, min(ws.max_column, 12) + 1)
                           if isinstance(ws.cell(r, c).value, str))
        if STOP.search(rowtext) and not any(not blank(ws.cell(r, c).value) and isinstance(ws.cell(r, c).value, (int, float))
                                            for c in vcols):
            break
        a = ws.cell(r, tb["label"])
        label = a.value if isinstance(a.value, str) else None
        if blank(label):
            alt = ws.cell(r, tb["alt"]).value
            label = alt if isinstance(alt, str) else None
        if blank(label):
            continue
        c0, c1, _, _ = span(ws, r, tb["label"])
        if c1 >= min(vcols):      # a heading that spans the value columns
            section_nls = bool(NLS.search(label))
            continue
        cells = {c: anchor(ws, r, c) for c in vcols}
        has_data = any(not blank(x.value) and not is_formula(x.value) for x in cells.values())
        code = waste_code(label, section_nls)
        m = mrows.get(code) if code else None
        if m is None:
            if has_data:
                for x in cells.values():
                    if not blank(x.value) and not is_formula(x.value):
                        x.value = None
                mark(next(iter(cells.values())))
                res.yellow += 1
                res.missing.append(re.sub(r"\s+", " ", label).strip()[:50])
                unmatched += 1
            continue
        used.add(code)
        vals = dict(m["values"])
        for c, k in vcols.items():
            x = cells[c]
            if is_formula(x.value):
                continue
            raw = nextport if k == "w_next_port" else vals.get(k)
            if k == "w_next_port" and not any(vals.get(q) for q in ("w_deliver", "w_onboard")):
                raw = None
            if blank(raw):
                if not blank(x.value):
                    x.value = None
                    cleared += 1
                continue
            x.value = raw
        filled += 1
    for code, m in mrows.items():
        if code not in used and m["values"]:
            res.note(f"Master waste line '{m['name']}' has no matching line on this form; not filled.")
    return filled, cleared, unmatched


def _label_ok(v):
    return isinstance(v, str) and bool(NUMLABEL.match(v))


def fill_fields(ctx, ws, res, upto_row):
    from forms import mark, has_border, anchor, span
    master = ctx["master"]
    f = master.fields
    wh = master.waste["header"]
    vals = {"ship_name": f.get("ship_name"), "imo": f.get("imo"), "gt": f.get("gt"), "nt": f.get("nt"),
            "ship_type": f.get("ship_type"), "owner": f.get("company_name"), "call_sign": f.get("call_sign"),
            "flag": f.get("flag"), "terminal": f.get("berth"), "eta_dt": f.get("eta_arrival") or f.get("eta"),
            "etd_dt": f.get("etd_departure"), "last_port": f.get("last_port"), "next_port": f.get("next_port"),
            "submitter": f.get("master_name"), "w_last_port": wh.get("w_last_port"), "w_last_date": wh.get("w_last_date"),
            "w_next_port": wh.get("w_next_port"), "company_imo": f.get("company_imo")}
    import msw as MSW
    msw = MSW.is_msw_book(ctx["wb"])
    if msw:
        def loc(name):
            for p in master.ports:
                if p.get("port") == name and p.get("locode"):
                    return p["locode"]
            return name
        vals["w_last_port"] = loc(vals["w_last_port"]) if vals["w_last_port"] else None
        vals["w_next_port"] = f.get("next_port_locode") or (loc(vals["w_next_port"]) if vals["w_next_port"] else None)
    seen, count = set(), 0
    for r in range(1, min(upto_row, ws.max_row + 1)):
        row_labels = [re.sub(r"\s+", " ", str(ws.cell(r, c).value)).lower() for c in range(1, min(ws.max_column, 30) + 1)
                      if isinstance(ws.cell(r, c).value, str)]
        for c in range(1, min(ws.max_column, 30) + 1):
            cell = ws.cell(r, c)
            v = cell.value
            if isinstance(cell, MergedCell) or (r, c) in seen or not isinstance(v, str) or not (3 < len(v.strip()) < 160):
                continue
            if MSW.TECH.match(v.strip()):
                continue
            lab = re.sub(r"^\d+(\.\d+)*\.?\s*", "", re.sub(r"\s+", " ", v).strip().lower())
            key = next((k for k, rx in FIELDS if re.search(rx, lab)), None)
            if re.search(r"port of last delivery", lab):
                key = "w_last_port"
            if msw and lab in ("date:", "date"):
                key = "w_last_date"
            if not key:
                continue
            if key == "imo" and any(re.search(r"compa[ñn][ií]a|\bcompany\b", x) for x in row_labels):
                key = "company_imo"
            c0, c1, r0, r1 = span(ws, r, c)
            below = anchor(ws, r1 + 1, c0)
            tgt = None

            def is_label(x):
                t = x.value
                return isinstance(t, str) and (_label_ok(t) or t.strip().endswith(":") or
                                               any(re.search(rx, re.sub(r"\s+", " ", t).lower()) for _, rx in FIELDS))
            if msw:
                rc = anchor(ws, r0, c1 + 1)
                if not blank(rc.value) and not is_label(rc):
                    tgt = rc
            if tgt is None and below.row < upto_row and not is_label(below) and (not blank(below.value) or has_border(below)) and below.row > r1:
                tgt = below
            for col in (range(c1 + 1, min(c1 + 7, ws.max_column + 1)) if tgt is None else []):
                cand = anchor(ws, r0, col)
                cv = cand.value
                if isinstance(cv, str) and cv.strip() and (any(re.search(rx, cv.lower()) for _, rx in FIELDS) or _label_ok(cv)):
                    break
                if not blank(cv) or has_border(cand):
                    tgt = cand
                    break
            if tgt is None:
                continue
            seen.add((tgt.row, tgt.column))
            val = vals.get(key)
            if blank(val):
                if not blank(tgt.value):
                    tgt.value = None          # old value of an earlier call
                mark(tgt); res.yellow += 1; res.missing.append(v.strip()[:45])
                continue
            if key in ("eta_dt", "etd_dt", "w_last_date"):
                d = C.to_date(val)
                if d:
                    keep_time = key != "w_last_date" and hasattr(val, "hour")
                    tgt.value = val if keep_time else d
                    if not C.is_date_format(tgt.number_format):
                        tgt.number_format = "dd-mm-yyyy hh:mm" if keep_time else "dd-mm-yyyy"
                    count += 1
                    continue
            tgt.value = val
            count += 1
    return count


def fill_waste_sheet(ctx, ws, res):
    """Returns True when this sheet was handled."""
    tb = find_table(ws)
    upto = tb["row"] if tb else ws.max_row
    n_head = fill_fields(ctx, ws, res, upto)
    if not tb and n_head < 3:
        return False
    res.kind = "Waste declaration"
    if tb:
        filled, cleared, unmatched = fill_table(ctx, ws, res, tb)
        line = f"{filled} waste lines filled"
        if unmatched:
            line += f", {unmatched} lines not in the master (old values removed)"
        res.lines.append(line)
        if cleared:
            res.note(f"{cleared} old value(s)/placeholder code(s) in the quantity cells were removed.")
    if n_head:
        res.lines.append(f"{n_head} header fields filled")
    return True

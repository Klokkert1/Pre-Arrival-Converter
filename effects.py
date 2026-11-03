"""Crew's effects as a long list: one line per person and item (family name, first names, rank, effect, quantity, unit)."""
import re
from openpyxl.cell.cell import MergedCell
from matching import norm, match
import vocabulary as V

ITEMS = [("eff_cig", "Cigarettes"), ("eff_tob", "Tobacco"), ("eff_wine", "Wine")]


def _find(ws):
    hdr = None
    for r in range(1, min(ws.max_row, 20) + 1):
        texts = [norm(str(ws.cell(r, c).value)) for c in range(1, min(ws.max_column, 14) + 1) if isinstance(ws.cell(r, c).value, str)]
        if any(re.search(r"\b(crew )?effects?\b|personal effects?", t) for t in texts) and any(t.startswith("quantity") for t in texts):
            hdr = r
    return hdr


def is_effects(ws):
    if not re.search(r"effect", ws.title, re.I) and not re.search(r"effect", " ".join(
            str(ws.cell(r, c).value) for r in range(1, 5) for c in range(1, 6) if ws.cell(r, c).value), re.I):
        return False
    return _find(ws) is not None


def _split(txt):
    m = re.match(r"^\s*([\d.,]+)\s*(.*)$", str(txt))
    if not m:
        return None, None
    try:
        q = float(m.group(1).replace(",", "."))
    except ValueError:
        return None, None
    return (int(q) if q.is_integer() else q), m.group(2).strip()


def fill_effects(ctx, ws, res):
    from forms import mark, anchor, get_allowed, is_formula, blank
    import convert as C
    r0 = _find(ws)
    cols = {}
    for c in range(1, min(ws.max_column, 14) + 1):
        v = ws.cell(r0, c).value
        if not isinstance(v, str):
            continue
        t = norm(v.lstrip("* "))
        if re.search(r"name", t) and match(v.lstrip("* "), V.PERSON)[0] in ("family_name", "given_names"):
            cols[match(v.lstrip("* "), V.PERSON)[0]] = c
        elif re.search(r"(seq|nr|no\b|number).*crew|crew.*(seq|nr|number)", t):
            cols["crewno"] = c
        elif re.search(r"\b(crew )?effects?\b|personal effect", t):
            cols["effect"] = c
        elif t.startswith("quantity"):
            cols["qty"] = c
        elif t.startswith("unit"):
            cols["unit"] = c
        elif t.startswith("description"):
            cols["desc"] = c
        elif re.match(r"^(seq|no$|nr$|serial)", t):
            cols["serial"] = c
        else:
            k = match(v.lstrip("* "), V.PERSON)[0]
            if k in ("family_name", "given_names", "rank"):
                cols[k] = c
    if "effect" not in cols or "qty" not in cols:
        return False
    lines = []
    for i, rec in enumerate(ctx["master"].crew, 1):
        for key, name in ITEMS:
            if rec.get(key):
                q, u = _split(rec[key])
                if q is None:
                    lines.append((i, rec, name, None, str(rec[key])))
                else:
                    lines.append((i, rec, name, q, u))
    res.kind = "Crew's effects"
    r = r0 + 1
    for n, (i, rec, name, q, u) in enumerate(lines, 1):
        row = r0 + n
        def put(key, val):
            if key in cols:
                cell = anchor(ws, row, cols[key])
                if not isinstance(cell, MergedCell) and not is_formula(cell.value):
                    cell.value = val
                return cell
        put("serial", n); put("crewno", i)
        put("family_name", rec.get("family_name")); put("given_names", rec.get("given_names")); put("rank", rec.get("rank"))
        put("effect", name); put("desc", None)
        cq = put("qty", q if q is not None else None)
        uc = put("unit", u or None)
        if q is None and cq is not None:
            mark(cq); res.yellow += 1
        if uc is not None and u:
            al = get_allowed(ctx["wb"], ws, uc)
            if al:
                kw = ("kilogram" if re.search(r"^(kg|kgs|kilo)", u, re.I) else
                      "litre" if re.search(r"^(l|ltr|litre|liter)", u, re.I) else
                      "number" if re.search(r"cig|bottle|piece|pcs|stick|pack|carton|unit|number", u, re.I) else None)
                pk = C.choose_from_list(u, al)
                if pk is None and kw:
                    pk = next((a for a in al if kw in str(a).lower()), None)
                if pk is None and kw == "number":
                    pk = next((a for a in al if re.search(r"piece|unit|article", str(a), re.I)), None)
                if pk is not None:
                    uc.value = pk
                else:
                    mark(uc); res.yellow += 1
    last = r0 + len(lines)
    cleared = 0
    for row in range(last + 1, min(ws.max_row, last + 400) + 1):
        for key, c in cols.items():
            if key == "serial":
                continue
            cell = ws.cell(row, c)
            if not isinstance(cell, MergedCell) and not blank(cell.value) and not is_formula(cell.value):
                cell.value = None; cleared += 1
    if cleared:
        res.note(f"{cleared} old value(s) from an earlier call were removed below the new lines.")
    res.lines.append(f"{len(lines)} effect lines filled")
    return True

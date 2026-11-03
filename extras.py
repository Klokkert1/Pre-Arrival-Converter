"""Extra header handling for two layouts: 'Ship pre-arrival security information' (ISPS) and the
Maritime Declaration of Health header. Values go to the cell RIGHT of the label (never over a label)."""
import re
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import PatternFill


def _txt(v):
    return re.sub(r"\s+", " ", str(v)).strip().lower()


def _sheet_text(ws, rows=25):
    out = []
    for row in ws.iter_rows(min_row=1, max_row=rows, max_col=20):
        for c in row:
            if isinstance(c.value, str):
                out.append(_txt(c.value))
    return " | ".join(out)


def kind(ws):
    t = _sheet_text(ws)
    if "pre-arrival security information" in t and "security level at which the ship" in t:
        return "isps"
    if "maritime declaration of health" in t and "net tonnage" in t:
        return "mdoh_head"
    return None


def _put(ws, cell, val, res, dateval=False):
    import convert as C
    from forms import anchor
    if cell is None or isinstance(cell, MergedCell) or val is None or str(val).strip() == "":
        return False
    if dateval:
        d = C.to_date(val)
        if not d:
            return False
        if not C.is_date_format(cell.number_format):
            cell.number_format = "dd-mm-yyyy"
        val = d
    cell.value = val
    if cell.fill is not None and cell.fill.fgColor is not None and cell.fill.fgColor.rgb == "FFFFFF00":
        cell.fill = PatternFill(fill_type=None)
    return True


def _right(ws, r, c):
    from forms import anchor, span
    c0, c1, r0, r1 = span(ws, r, c)
    return anchor(ws, r0, c1 + 1)


def post(ctx, ws, res):
    k = kind(ws)
    if not k:
        return
    from forms import anchor
    f = ctx["master"].fields
    dep = str(f.get("movement") or "Arrival").strip().lower().startswith("dep")
    n = 0
    for r in range(1, 40):
        for c in range(1, 20):
            cell = ws.cell(r, c)
            if isinstance(cell, MergedCell) or not isinstance(cell.value, str):
                continue
            t = _txt(cell.value)
            tgt = _right(ws, r, c)
            val, isdate = None, False
            if k == "isps":
                if t.startswith("name of company"):
                    val = " / ".join(str(x) for x in (f.get("company_name") or f.get("owner"), f.get("company_imo")) if x)
                elif t.startswith("cso name"):
                    nm = " ".join(str(x) for x in (f.get("cso_given"), f.get("cso_family")) if x)
                    val = ", ".join(x for x in (nm, f.get("cso_phone")) if x)
                elif t.startswith("inmarsat"):
                    val = f.get("inmarsat")
                elif t.startswith("port facility of arrival"):
                    val = f.get("port_facility")
                elif t.startswith("primary purpose of call"):
                    val = f.get("call_purpose")
                elif t.startswith("location of ship"):
                    val = f.get("ship_location")
                elif t.startswith("does the ship have a valid international"):
                    # same row: value, certificate number, issued by, expiry (header row above)
                    _isps_row(ws, r, f, res)
                    continue
                elif t.startswith("does the ship have an approved ssp"):
                    tgt = _right(ws, r, 4)
                    val = f.get("ssp_onboard")
                elif t.startswith("security level at which"):
                    val = f.get("ship_sec_level")
            else:  # mdoh header
                if t.startswith("net tonnage"):
                    val = f.get("nt")
                elif t.rstrip(":") == "from":
                    val = f.get("port") if dep else f.get("last_port")
                elif t.rstrip(":") == "to":
                    val = f.get("next_port") if dep else f.get("port")
                elif t == "dated":
                    val, isdate = f.get("san_date"), True
                elif t.startswith("issued at"):
                    val = f.get("san_place")
            if val is not None and _put(ws, tgt, val, res, isdate):
                n += 1
    if k == "mdoh_head":
        t = ctx["master"].fields.get("san_type")
        if t:
            res.note(f"Sanitation certificate: the master says '{t}' - please tick Certificate or Exemption by hand.")
    res.lines.append(f"{n} extra header fields filled")


def _isps_row(ws, r, f, res):
    from forms import anchor
    # header texts on the row above give the columns
    cols = {}
    for c in range(1, 25):
        v = ws.cell(r - 1, c).value
        if isinstance(v, str):
            t = _txt(v)
            if t.startswith("certificate number"): cols["no"] = c
            elif t.startswith("issued by"): cols["by"] = c
            elif t.startswith("expiry"): cols["exp"] = c
    from forms import span
    # yes/no cell: first value cell after the label block
    lab_end = span(ws, r, 1)[1]
    _put(ws, anchor(ws, r, lab_end + 1), f.get("issc_valid"), res)
    if "no" in cols: _put(ws, anchor(ws, r, cols["no"]), f.get("issc_number"), res)
    if "by" in cols: _put(ws, anchor(ws, r, cols["by"]), f.get("issc_issuer"), res)
    if "exp" in cols: _put(ws, anchor(ws, r, cols["exp"]), f.get("issc_expiry"), res, True)


# ---------------------------------------------------------------- report: labels that were not filled
WRITES = set()
_ORIG = {}


def start_tracking():
    """Remember every cell the program writes to (even when it writes the same value that was already there)."""
    from openpyxl.cell.cell import Cell
    WRITES.clear()
    if "p" in _ORIG:
        return
    p = Cell.__dict__["value"]
    _ORIG["p"] = p

    def setter(self, v):
        try:
            WRITES.add((self.parent.title, self.row, self.column))
        except Exception:
            pass
        p.fset(self, v)
    Cell.value = property(p.fget, setter)


def stop_tracking():
    from openpyxl.cell.cell import Cell
    if "p" in _ORIG:
        Cell.value = _ORIG.pop("p")


def snapshot_values(wb):
    """All filled-in values of the form as it came in (to find old values that nobody replaced)."""
    snap = {}
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        snap[ws.title] = {(r, c): cell.value for (r, c), cell in list(ws._cells.items())
                          if r <= 60 and c <= 20 and cell.value not in (None, "")
                          and not (isinstance(cell.value, str) and cell.value.startswith("="))}
    return snap


def report_old_values(ws, res, labels, values):
    """A label that stayed as it was, with a value from an earlier call that nobody replaced: yellow + note."""
    from forms import anchor, span, mark
    import vocabulary as V
    from matching import match

    def is_label(t):
        t = str(t).strip()
        if re.match(r"^\d+(\.\d+)?\.?\s+[A-Za-z(]", t) and len(t) > 5:
            return True
        if t.endswith(":"):
            return True
        return any(match(t, v)[0] for v in (V.FIELDS, V.PERSON, V.PORTS, V.WASTE))

    real = {k: t for k, t in labels.items() if is_label(t)}
    row_count = {}
    for (r, c) in real:
        row_count[r] = row_count.get(r, 0) + 1
    flagged = []
    for (r, c), text in sorted(real.items()):
        cell = ws.cell(r, c)
        if isinstance(cell, MergedCell) or cell.value != text or row_count.get(r, 0) >= 4:
            continue
        low = text.lower()
        if len(text) < 4 or text.endswith("?") or "_" in text or (text.isupper() and not text.endswith(":")) or re.search(
                r"signature|signed|stamp|convention|certify|declaration of|^arrival$|^departure$|^yes$|^no$|remarks|attached", low):
            continue
        try:
            c0, c1, r0, r1 = span(ws, r, c)
            below = anchor(ws, r1 + 1, c0)
            right = anchor(ws, r0, c1 + 1)
        except Exception:
            continue
        numbered = bool(re.match(r"^\d+(\.\d+)?\.?\s+\S", text))
        target = None
        for cand in ((below, right) if numbered else (right, below)):
            if cand is None or isinstance(cand, MergedCell):
                continue
            key = (cand.row, cand.column)
            if key in real or key not in values or cand.value != values[key] or (ws.title, cand.row, cand.column) in WRITES:
                continue                  # not an old value, or already replaced
            if isinstance(cand.value, str) and (cand.value.strip().endswith(":") or len(cand.value) > 80):
                continue
            if cand.fill is not None and cand.fill.fgColor is not None and cand.fill.fgColor.rgb == "FFFFFF00":
                continue
            target = cand
            break
        if target is None:
            continue
        group = [target]
        cc = target.column + 1
        while cc <= 20 and cc <= target.column + 3:        # further old values in the same row of boxes
            nxt = ws.cell(target.row, cc)
            if isinstance(nxt, MergedCell):
                cc += 1
                continue
            if (nxt.row, nxt.column) in real or (nxt.row, nxt.column) not in values or nxt.value != values[(nxt.row, nxt.column)] \
                    or (ws.title, nxt.row, nxt.column) in WRITES:
                break
            group.append(nxt)
            cc += 1
        for g in group:
            mark(g)
        res.yellow += len(group)
        flagged.append(re.sub(r"\s+", " ", text).rstrip(": ")[:40])
    if flagged:
        shown = "; ".join(flagged[:10]) + (f" (+{len(flagged) - 10} more)" if len(flagged) > 10 else "")
        res.note(f"Old values from an earlier call were left in place (yellow) because the master has nothing for them: "
                 f"{shown}. Please check or replace them.")


def snapshot_labels(wb):
    """Text cells of the form as it came in (these are the form's own labels)."""
    snap = {}
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        d = {}
        for (r, c), cell in list(ws._cells.items()):
            if r <= 40 and c <= 20 and isinstance(cell.value, str) and 3 <= len(cell.value.strip()) <= 70 \
                    and not cell.value.startswith("="):
                d[(r, c)] = cell.value.strip()
        snap[ws.title] = d
    return snap


def report_unfilled(ws, res, labels):
    """Empty boxes with a form label on their left that nobody filled and that are not marked yellow."""
    from forms import has_border, anchor
    out = []
    for (r, c), text in sorted(labels.items()):
        cell = ws.cell(r, c)
        if isinstance(cell, MergedCell) or cell.value != text:
            continue
        right = None
        try:
            from forms import span
            c0, c1, r0, r1 = span(ws, r, c)
            right = anchor(ws, r0, c1 + 1)
        except Exception:
            continue
        if right is None or isinstance(right, MergedCell):
            continue
        below = anchor(ws, r1 + 1, c0)
        if below is not None and not isinstance(below, MergedCell) and below.value not in (None, "") \
                and labels.get((below.row, below.column)) != str(below.value).strip():
            continue                       # the value sits below the label
        if right.value not in (None, ""):
            continue
        if right.fill is not None and right.fill.fgColor is not None and right.fill.fgColor.rgb == "FFFFFF00":
            continue                       # already reported as missing/yellow
        if not has_border(right):
            continue
        if (r, c + 1) in labels:
            continue
        if any(ws.cell(r, cc).value not in (None, "") and not isinstance(ws.cell(r, cc), MergedCell)
               and labels.get((r, cc)) != str(ws.cell(r, cc).value).strip() for cc in range(c1 + 1, min(c1 + 4, 21))):
            continue                       # a value further right in the same row
        low = text.lower()
        if "_" in text or re.search(r"signature|signed|stamp|sign\b|convention|certify|declaration of|supplemental|^arrival$|^departure$|^yes$|^no$|\(l\)|\(kg\)|\(pcs\)", low):
            continue
        if re.match(r"^\d+ ?(till|until|to|-)", low):
            continue
        if text.isupper() and not text.endswith(":"):
            continue
        out.append(re.sub(r"\s+", " ", text).rstrip(": ")[:45])
    if out:
        shown = "; ".join(out[:10]) + (f" (+{len(out) - 10} more)" if len(out) > 10 else "")
        res.note(f"Boxes left empty because the label was not recognised: {shown}. "
                 f"Fill these by hand - or send me this form so the word list can be extended.")

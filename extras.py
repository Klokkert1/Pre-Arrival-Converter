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

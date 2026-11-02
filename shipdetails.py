"""'Ship details' sheets with label (ending in ':') in one column and the value to its right (Kiel Canal / German ports)."""
import re
from openpyxl.cell.cell import MergedCell
from matching import norm

SPEC = [
    (r"^ship name", "ship_name"), (r"^call ?sign", "call_sign"), (r"^imo number", "imo"), (r"^mmsi", "mmsi"),
    (r"^eni number", None), (r"^flag", "flag"), (r"^length overall", "length"), (r"^beam", "breadth"),
    (r"^gross tonnage", "gt"), (r"^net tonnage", "nt"), (r"^port of regist", "port_registry"),
    (r"^ship type|^type of ship", "ship_type"), (r"^ships? phone", "ship_phone"), (r"^deadweight|^dwt", "dwt"),
    (r"^next port", "next_port"), (r"^ism name of company|^name of company", "company_name"),
    (r"^ism company id", "company_imo"), (r"^draught forward|^draft forward", "draft_fore"),
    (r"^draught aft|^draft aft", "draft_aft"), (r"^ships? e ?mail", "ship_email"),
]


def is_shipdetails(ws):
    n = 0
    for r in range(1, 45):
        v = ws.cell(r, 2).value
        if isinstance(v, str) and v.strip().endswith(":") and any(re.search(rx, norm(v)) for rx, _ in SPEC):
            n += 1
    return n >= 6


def fill_shipdetails(ctx, ws, res):
    from forms import mark, anchor, span, is_formula, blank
    f = ctx["master"].fields
    n = 0
    for r in range(1, min(ws.max_row, 60) + 1):
        for c in range(1, 4):
            cell = ws.cell(r, c)
            v = cell.value
            if isinstance(cell, MergedCell) or not isinstance(v, str) or not v.strip().endswith(":"):
                continue
            key = next(((k) for rx, k in SPEC if re.search(rx, norm(v))), "?")
            if key == "?":
                continue
            c0, c1, r0, r1 = span(ws, r, c)
            tgt = anchor(ws, r0, c1 + 1)
            if tgt is None or isinstance(tgt, MergedCell) or is_formula(tgt.value):
                continue
            val = f.get(key) if key else None
            if key == "next_port" and val is None:
                continue
            if blank(val):
                if not blank(tgt.value):
                    tgt.value = None
                if "*" in v:
                    mark(tgt); res.yellow += 1; res.missing.append(v.strip(" :*")[:40])
                continue
            tgt.value = val
            n += 1
    res.kind = "Ship details"
    res.lines.append(f"{n} fields filled")
    return n > 0

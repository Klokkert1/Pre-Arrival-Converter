"""IMO FAL 1 (General Declaration): numbered labels with the value in the cell below."""
import re
from openpyxl.cell.cell import MergedCell
import convert as C

SPEC = [  # (regex on the label without its number, key)
    (r"name and description of ship", "name_desc"),
    (r"port of arrival", "port_arr"),
    (r"date.{0,12}arrival|date and time of arrival", "date_arr"),
    (r"nationality of ship", "flag"),
    (r"name of master", "master"),
    (r"port of departure", "port_dep"),
    (r"certificate of registry", "cert"),
    (r"name and address of ship.?s agent|ship.?s agent", "agent"),
    (r"gross tonnage", "gt"),
    (r"net tonnage", "nt"),
    (r"position of the ship in the port", "position"),
    (r"brief particulars of voyage", "particulars"),
    (r"brief description of the cargo", "cargo"),
    (r"number of crew", "n_crew"),
    (r"number of passengers", "n_pax"),
    (r"remarks", "remarks"),
    (r"cargo declaration", "att_cargo"),
    (r"ship.?s stores declaration", "att_stores"),
    (r"crew list", "att_crew"),
    (r"passenger list", "att_pax"),
    (r"crew.?s effects declaration", "att_effects"),
]
NUM = re.compile(r"^\s*\d{1,2}\s*\.\s*")


def is_fal1(ws):
    n = 0
    for r in range(1, 14):
        for c in range(1, 12):
            v = ws.cell(r, c).value
            if isinstance(v, str) and re.search(r"name and description of ship|port of arrival|nationality of ship|general declaration", v, re.I):
                n += 1
    return n >= 3


def fill_fal1(ctx, ws, res):
    from forms import mark, anchor, span, is_formula, blank
    m = ctx["master"]
    f = m.fields
    mv = str(f.get("movement") or "Arrival").lower()
    dep = mv.startswith("dep")
    port = f.get("port")
    nd = " / ".join(str(x) for x in (f.get("ship_name"), f.get("ship_type")) if x) or None
    parts = f.get("voyage_particulars") or (f"{f.get('last_port')} - {f.get('port')} - {f.get('next_port')}"
                                            if f.get("last_port") and f.get("next_port") else None)
    vals = {
        "name_desc": nd, "port_arr": port if not dep else f.get("last_port"),
        "date_arr": f.get("eta_arrival") or f.get("eta"), "flag": f.get("flag"), "master": f.get("master_name"),
        "port_dep": f.get("next_port") if not dep else port, "cert": f.get("cert_registry"), "agent": f.get("agent"),
        "gt": f.get("gt"), "nt": f.get("nt"), "position": f.get("ship_location") or f.get("berth"),
        "particulars": parts, "cargo": f.get("cargo_desc"),
        "n_crew": f.get("n_crew") if f.get("n_crew") not in (None, "") else (len(m.crew) or None),
        "n_pax": f.get("n_pax") if f.get("n_pax") not in (None, "") else (len(m.pax) if m.pax else (0 if m.crew else None)),
        "remarks": f.get("remarks"), "att_cargo": f.get("att_cargo"), "att_stores": f.get("att_stores"),
        "att_crew": f.get("att_crew"), "att_pax": f.get("att_pax"), "att_effects": f.get("att_effects"),
    }
    n_ok, skipped = 0, 0
    for r in range(1, min(ws.max_row, 60) + 1):
        for c in range(1, min(ws.max_column, 16) + 1):
            cell = ws.cell(r, c)
            v = cell.value
            if isinstance(cell, MergedCell) or not isinstance(v, str) or not NUM.match(v):
                continue
            lab = NUM.sub("", re.sub(r"\s+", " ", v).strip().lower())
            key = next((k for rx, k in SPEC if re.search(rx, lab)), None)
            if not key:
                continue
            c0, c1, r0, r1 = span(ws, r, c)
            tgt = anchor(ws, r1 + 1, c0)
            if isinstance(tgt.value, str) and NUM.match(tgt.value):
                continue
            if is_formula(tgt.value):
                skipped += 1
                continue
            val = vals.get(key)
            if blank(val) and key.startswith("att_"):
                continue                      # standard text such as 'Attached' stays as the company has it
            if blank(val):
                if not blank(tgt.value):
                    tgt.value = None
                mark(tgt); res.yellow += 1; res.missing.append(v.strip()[:45])
                continue
            if key == "date_arr":
                d = C.to_date(val)
                if d:
                    tgt.value = d
                    if not C.is_date_format(tgt.number_format):
                        tgt.number_format = "dd-mm-yyyy"
                    n_ok += 1
                    continue
            tgt.value = val
            n_ok += 1
    res.kind = "FAL 1 General Declaration"
    res.lines.append(f"{n_ok} fields filled")
    if skipped:
        res.note(f"{skipped} field(s) are filled by formulas from other sheets; left untouched.")
    return n_ok > 0 or skipped > 0

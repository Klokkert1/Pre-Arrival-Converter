"""UK NMSW FAL 1 / FAL 2 / FAL 7: label cell (with help text) and the value in the cell to its right."""
import re, datetime
from openpyxl.cell.cell import MergedCell
from matching import norm
import convert as C

SPEC = [  # (regex on normalised first words of the label, key)
    (r"^arrival departure", "movement"), (r"^name of ship agent", "agent"), (r"^contact number of ship agent", "agent_phone"), (r"^name of ship$|^name of ship ", "ship_name"), (r"^imo number", "imo"),
    (r"^call sign", "call_sign"), (r"^name of master", "master"), (r"^flag state", "flag"),
    (r"^(un )?locode of arrival port|^arrival port locode", "arr_port"),
    (r"^(un )?locode of departure port|^departure port locode", "dep_port"),
    (r"^date of arrival", "arr_date"), (r"^time of arrival", "arr_time"),
    (r"^date of departure", "dep_date"), (r"^time of departure", "dep_time"),
    (r"^voyage number", "voyage"), (r"^last port of call", "last_port"), (r"^next port of call", "next_port"),
    (r"^gross tonnage", "gt"), (r"^net tonnage", "nt"),
    (r"^certificate of registry", "cert"), (r"^position of ship", "position"),
    (r"^brief particulars", "particulars"), (r"^brief description of the cargo", "cargo"),
    (r"^number of crew", "n_crew"), (r"^number of passengers", "n_pax"), (r"^remarks", "remarks"),
]
KEEP = {"agent"}      # master has the agent as one text; contact number not separate -> handled below


def is_nmsw(ws):
    t = " ".join(str(ws.cell(r, c).value) for r in range(1, 3) for c in range(1, 4) if ws.cell(r, c).value)
    return bool(re.search(r"nmsw", t, re.I)) and bool(re.search(r"fal ?[127]", t, re.I))


def lo(v):
    return re.sub(r"[^A-Z0-9]", "", str(v).upper()) if v else v


def fill_nmsw(ctx, ws, res):
    from forms import mark, anchor, span, is_formula, blank, get_allowed
    m = ctx["master"]
    f = m.fields
    dep = str(f.get("movement") or "Arrival").lower().startswith("dep")
    eta = f.get("eta_arrival") or f.get("eta")
    etd = f.get("etd_departure")
    port = lo(f.get("port_locode")) or None
    vals = {
        "ship_name": f.get("ship_name"), "imo": f.get("imo"), "call_sign": f.get("call_sign"),
        "master": f.get("master_name"), "flag": f.get("flag"), "agent_phone": f.get("agent_phone"), "voyage": f.get("voyage_no"), "gt": f.get("gt"), "nt": f.get("nt"),
        "arr_port": port if not dep else lo(f.get("next_port_locode")),
        "dep_port": lo(f.get("last_port_locode")) if not dep else port,
        "last_port": lo(f.get("last_port_locode")), "next_port": lo(f.get("next_port_locode")),
        "agent": f.get("agent"), "cert": f.get("cert_registry"), "position": f.get("ship_location") or f.get("berth"),
        "particulars": f.get("voyage_particulars"), "cargo": f.get("cargo_desc"),
        "n_crew": f.get("n_crew") if f.get("n_crew") not in (None, "") else (len(m.crew) or None),
        "n_pax": f.get("n_pax") if f.get("n_pax") not in (None, "") else (len(m.pax) if m.pax else (0 if m.crew else None)),
        "remarks": f.get("remarks"),
        "arr_date": eta if not dep else None, "dep_date": etd if not dep else (etd or eta),
        "arr_time": eta if not dep else None, "dep_time": etd if not dep else (etd or eta),
    }
    n = 0
    for r in range(1, min(ws.max_row, 40) + 1):
        for c in range(1, min(ws.max_column, 14) + 1):
            cell = ws.cell(r, c)
            v = cell.value
            if isinstance(cell, MergedCell) or not isinstance(v, str):
                continue
            first = norm(v.split("\n")[0])
            first = re.sub(r"\b(required|for example.*)$", "", first).strip()
            key = next((k for rx, k in SPEC if re.search(rx, first)), None)
            if not key:
                continue
            c0, c1, r0, r1 = span(ws, r, c)
            tgt = anchor(ws, r0, c1 + 1)
            if tgt is None or isinstance(tgt, MergedCell) or is_formula(tgt.value):
                continue
            if isinstance(tgt.value, str) and tgt.value.strip() and any(re.search(rx, norm(tgt.value.split("\n")[0])) for rx, _ in SPEC):
                continue                          # right neighbour is itself a label
            val = vals.get(key)
            if key == "movement":
                allowed = get_allowed(ctx["wb"], ws, tgt) or []
                want = "depart" if dep else "arriv"
                val = next((a for a in allowed if want in a.lower()), "Departure" if dep else "Arrival")
            elif key == "flag" and val:
                allowed = get_allowed(ctx["wb"], ws, tgt)
                v2, ok = C.convert_country(val, allowed, "iso 3 alpha-3")
                val = v2 if ok and v2 else val
            if blank(val):
                if not blank(tgt.value):
                    tgt.value = None
                mark(tgt); res.yellow += 1; res.missing.append(first[:40])
                continue
            if key in ("arr_date", "dep_date"):
                d = C.to_date(val)
                if not d:
                    continue
                tgt.value = d
                tgt.number_format = "dd/mm/yyyy"
            elif key in ("arr_time", "dep_time"):
                if not hasattr(val, "hour"):
                    mark(tgt); res.yellow += 1; res.missing.append(first[:40])
                    if not blank(tgt.value):
                        tgt.value = None
                    continue
                tgt.value = datetime.time(val.hour, val.minute)
                tgt.number_format = "hh:mm"
            else:
                tgt.value = val
            n += 1
    # 'Name of Master ... Completing Form' at the bottom of FAL 2 / 7
    for r in range(10, min(ws.max_row, 60) + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and norm(v).startswith("name of master authorised agent or officer completing"):
            tgt = ws.cell(r, 2)
            if not isinstance(tgt, MergedCell) and not is_formula(tgt.value) and f.get("master_name"):
                tgt.value = f.get("master_name")
                n += 1
    res.kind = "UK NMSW form"
    res.lines.append(f"{n} fields filled")
    res.note("Cargo / dangerous goods lines are not filled by this version; please complete them.") \
        if re.search(r"fal ?[27]", ws.title, re.I) else None
    return n > 0

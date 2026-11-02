"""Ports of Call list (company format): ship particulars, ISSC data and a table of the last ports."""
import re
from openpyxl.cell.cell import MergedCell
from matching import norm
import convert as C


def find_table(ws):
    for r in range(1, min(ws.max_row, 40) + 1):
        cols = {}
        for c in range(1, min(ws.max_column, 14) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str):
                t = norm(v)
                for k, rx in (("fac", r"port facility"), ("din", r"^date in$|^date from"), ("dout", r"^date out$|^date to"),
                              ("ship", r"^ship level|^ship security"), ("port", r"^port level|^port security"),
                              ("dos", r"^d o s$|^dos$|declaration of security"), ("rem", r"^remarks?$")):
                    if re.search(rx, t) and k not in cols:
                        cols[k] = c
                        break
        if {"din", "dout", "dos"} <= set(cols) or ({"fac", "din", "dout"} <= set(cols) and "port" in cols):
            return r, cols
    return None


def is_poc(ws):
    for r in range(1, 16):
        for c in range(1, 9):
            v = ws.cell(r, c).value
            if isinstance(v, str) and re.search(r"ports of call list", v, re.I):
                return find_table(ws) is not None
    return False


def extra_port_cols(master):
    """Per-port values from the master's Last_10_Ports sheet that the generic port records do not keep."""
    out = []
    if "Last_10_Ports" not in master.wb.sheetnames:
        return out
    ws = master.wb["Last_10_Ports"]
    hdr = {}
    for c in range(1, ws.max_column + 1):
        h = norm(ws.cell(5, c).value or "")
        if h.startswith("port security"):
            hdr["port"] = c
        elif h.startswith("declaration of security"):
            hdr["dos"] = c
        elif h.startswith("ship security"):
            hdr["ship"] = c
        elif h == "port" or h.startswith("port name"):
            hdr.setdefault("name", c)
    for r in range(6, 16):
        if ws.cell(r, hdr.get("name", 2)).value in (None, ""):
            continue
        out.append({k: ws.cell(r, c).value for k, c in hdr.items() if k != "name"})
    return out


def fill_poc(ctx, ws, res):
    from forms import mark, anchor, span, is_formula, blank
    m = ctx["master"]
    f = m.fields
    n = 0

    def put(cell, val):
        nonlocal n
        if isinstance(cell, MergedCell) or is_formula(cell.value):
            return
        if blank(val):
            if not blank(cell.value):
                cell.value = None
            mark(cell); res.yellow += 1
            return
        cell.value = val
        n += 1
    hit = find_table(ws)
    hr, cols = hit
    # label rows above the table
    for r in range(1, hr):
        for c in range(1, min(ws.max_column, 10) + 1):
            cell = ws.cell(r, c)
            v = cell.value
            if isinstance(cell, MergedCell) or not isinstance(v, str):
                continue
            t = norm(v)
            c0, c1, r0, r1 = span(ws, r, c)
            below = anchor(ws, r1 + 1, c0)
            right = anchor(ws, r0, c1 + 1)
            from forms import has_border
            for cc in range(c1 + 1, min(c1 + 5, ws.max_column + 1)):
                cand = anchor(ws, r0, cc)
                sp = span(ws, r0, cc)
                if sp[1] > sp[0]:
                    right = cand
                    break
            if t == "ships name":
                put(below, f.get("ship_name"))
            elif t == "official number":
                imo_prefix = isinstance(below.value, str) and below.value.strip().lower().startswith("imo")
                put(below, f"IMO {f.get('imo')}" if imo_prefix and f.get("imo") else f.get("official_no"))
            elif t == "port of registry":
                put(below, f.get("port_registry"))
            elif t == "flag":
                put(below, f.get("flag"))
            elif t in ("no", "number") and r > 8:
                put(right, f.get("issc_number"))
            elif t.startswith("issue date"):
                d = C.to_date(f.get("issc_issued"))
                put(right, d)
                if d is not None and not isinstance(right, MergedCell):
                    right.number_format = "dd-mm-yyyy"
            elif t.startswith("expiry date"):
                d = C.to_date(f.get("issc_expiry"))
                put(right, d)
                if d is not None and not isinstance(right, MergedCell):
                    right.number_format = "dd-mm-yyyy"
    extra = extra_port_cols(m)
    ports = m.ports[:10]
    for i in range(10):
        row = hr + 1 + i
        if i < len(ports):
            p = ports[i]
            e = extra[i] if i < len(extra) else {}
            put(ws.cell(row, cols["din"]), C.to_date(p.get("arrival")))
            put(ws.cell(row, cols["dout"]), C.to_date(p.get("departure")))
            for k, key in (("fac", None),):
                pass
            if "port" in cols:
                put(ws.cell(row, cols["port"]), e.get("port"))
            if "ship" in cols:
                put(ws.cell(row, cols["ship"]), e.get("ship") or f.get("ship_sec_level"))
            if "dos" in cols:
                put(ws.cell(row, cols["dos"]), e.get("dos"))
            if "fac" in cols and not is_formula(ws.cell(row, cols["fac"]).value):
                name = " ".join(str(x) for x in (p.get("port"), p.get("facility")) if x)
                put(ws.cell(row, cols["fac"]), name or None)
            for k in ("din", "dout"):
                cell = ws.cell(row, cols[k])
                if not isinstance(cell, MergedCell) and not is_formula(cell.value) and cell.value is not None:
                    cell.number_format = "dd-mm-yyyy"
        else:
            for k in ("din", "dout", "port", "dos"):
                if k in cols:
                    cell = ws.cell(row, cols[k])
                    if not isinstance(cell, MergedCell) and not is_formula(cell.value):
                        cell.value = None
    res.kind = "Ports of Call list"
    res.lines.append(f"{n} values filled, {len(ports)} port calls")
    return True

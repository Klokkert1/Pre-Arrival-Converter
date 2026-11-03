"""UK CERS3 / PortPlus workbook: sheets Vessel, Voyage, Waste, Security with label-left / value-right blocks."""
import re, datetime
from openpyxl.cell.cell import MergedCell
from matching import norm
import convert as C

SPEC = [
    (r"^imo number$", "imo"), (r"^ship name$", "ship_name"), (r"^call sign$", "call_sign"), (r"^mmsi number$", "mmsi"),
    (r"^flag state$", "flag"), (r"^gross tonnage$", "gt"), (r"^ship type$", "ship_type"),
    (r"^select port of registry$|^port of registry name$", "port_registry"),
    (r"^company name$", "company_name"), (r"^company imo number$", "company_imo"),
    (r"^inmarsat call number 1$", "inmarsat"),
    (r"^select port of call$|^port of call name$", "port"), (r"^port of call locode$", "port_locode"),
    (r"^port facility$", "port_facility"), (r"^position in port of call$", "position"),
    (r"^eta to port of call$", "eta"), (r"^etd from port of call$", "etd"),
    (r"^select last port$|^last port name$", "last_port"), (r"^last port locode$", "last_port_locode"),
    (r"^select next port$|^next port name$", "next_port"), (r"^next port locode$", "next_port_locode"),
    (r"^brief cargo description$", "cargo_desc"), (r"^persons on board toward port of call$", "persons"),
    (r"^last port delivered locode$", "w_last_locode"), (r"^last port delivered name$", "w_last_port"),
    (r"^last port delivered date$", "w_last_date"), (r"^port locode$", "port_locode"), (r"^port name$", "port"),
    (r"^location name$", "position"),
    (r"^first name$", "cso_given"), (r"^last name$", "cso_family"), (r"^phone$", "cso_phone"), (r"^email$", "cso_email"),
    (r"^issc number$", "issc_number"), (r"^issuer$", "issc_issuer"), (r"^issuer type$", "issc_issuer_type"),
    (r"^expiry$", "issc_expiry"),
]
YELLOW_KEYS = {"imo", "ship_name", "call_sign", "flag", "gt", "port", "port_locode", "eta"}
DATES = {"eta", "etd", "w_last_date", "issc_expiry"}


def is_cers(wb):
    names = set(wb.sheetnames)
    if not {"Vessel", "Voyage", "Waste"} <= names or "Summary" not in names:
        return False
    return "cers" in str(wb["Summary"]["B2"].value or "").lower()


def _vals(m):
    f = m.fields
    ports = m.ports
    last = ports[0] if ports else {}
    return {
        "imo": f.get("imo"), "ship_name": f.get("ship_name"), "call_sign": f.get("call_sign"), "mmsi": f.get("mmsi"),
        "flag": f.get("flag"), "gt": f.get("gt"), "ship_type": f.get("ship_type"), "port_registry": f.get("port_registry"),
        "company_name": f.get("company_name"), "company_imo": f.get("company_imo"), "inmarsat": f.get("inmarsat"),
        "port": f.get("port"), "port_locode": f.get("port_locode"), "port_facility": f.get("port_facility"),
        "position": f.get("ship_location") or f.get("berth"),
        "eta": f.get("eta_arrival") or f.get("eta"), "etd": f.get("etd_departure"),
        "last_port": f.get("last_port"), "last_port_locode": f.get("last_port_locode") or last.get("locode"),
        "next_port": f.get("next_port"), "next_port_locode": f.get("next_port_locode"),
        "cargo_desc": f.get("cargo_desc"),
        "persons": (len(m.crew) + len(m.pax)) or None,
        "w_last_port": m.waste["header"].get("w_last_port"), "w_last_date": m.waste["header"].get("w_last_date"),
        "w_last_locode": None,
        "cso_given": f.get("cso_given"), "cso_family": f.get("cso_family"), "cso_phone": f.get("cso_phone"),
        "cso_email": f.get("cso_email"), "issc_number": f.get("issc_number"), "issc_issuer": f.get("issc_issuer"),
        "issc_issuer_type": f.get("issc_issuer_type"), "issc_expiry": f.get("issc_expiry"),
    }


def fill_cers_header(ctx, ws, res):
    from forms import mark, anchor, span, is_formula, blank, has_border
    m = ctx["master"]
    vals = _vals(m)
    if not vals["w_last_locode"]:
        # LOCODE of the last waste-delivery port, when the port is one of the master's last ports
        for p in m.ports:
            if p.get("port") and p.get("port") == vals["w_last_port"]:
                vals["w_last_locode"] = p.get("locode")
    sheet = ws.title.lower()
    cso_block = sheet == "security"
    n = 0
    seen = set()
    for r in range(1, min(ws.max_row, 40) + 1):
        for c in range(1, min(ws.max_column, 16) + 1):
            cell = ws.cell(r, c)
            v = cell.value
            if isinstance(cell, MergedCell) or not isinstance(v, str) or v.startswith("="):
                continue
            t = norm(v)
            key = next((k for rx, k in SPEC if re.search(rx, t)), None)
            if not key:
                continue
            # CSO / agent / ISSC sit in the right-hand block of the Security sheet; the agent block shares the labels
            if cso_block and key in ("cso_phone", "cso_email", "cso_given", "cso_family"):
                if r > 11 and key in ("cso_phone", "cso_email"):
                    key = {"cso_phone": "agent_phone", "cso_email": "agent_email"}[key]
            if key in ("issc_expiry",) and not cso_block:
                continue
            if cso_block and (r >= 33 or key in ("port", "port_locode", "position")):
                continue
            if key in ("port", "port_locode", "position") and sheet == "waste" and r < 6:
                continue
            c0, c1, r0, r1 = span(ws, r, c)
            tgt = None
            for boxed in (True, False):          # first look for the input box (left border), then any bordered/filled cell
                for cc in range(c1 + 1, min(c1 + 5, ws.max_column + 1)):
                    cand = anchor(ws, r0, cc)
                    cv = cand.value
                    if isinstance(cv, str) and cv.strip() and any(re.search(rx, norm(cv)) for rx, _ in SPEC):
                        break
                    if isinstance(cv, str) and len(cv.split()) >= 7 and len(cv) > 45:
                        break                      # a note or instruction, not an input box
                    if isinstance(cv, str) and cv.startswith("="):
                        break
                    if boxed:
                        if cand.border.left is not None and cand.border.left.style:
                            tgt = cand
                            break
                    elif not blank(cv) or has_border(cand):
                        tgt = cand
                        break
                if tgt is not None:
                    break
            if tgt is None or isinstance(tgt, MergedCell) or (tgt.row, tgt.column) in seen:
                continue
            seen.add((tgt.row, tgt.column))
            val = vals.get(key) if key in vals else {"agent_phone": None, "agent_email": None}.get(key)
            if key == "agent_phone" or key == "agent_email":
                val = None
            if blank(val):
                had = not blank(tgt.value)
                if had:
                    tgt.value = None
                if had or key in YELLOW_KEYS:
                    mark(tgt); res.yellow += 1; res.missing.append(v.strip()[:40])
                continue
            if key in DATES:
                if key in ("eta", "etd") and hasattr(val, "hour"):
                    tgt.value = val
                    tgt.number_format = "dd/mm/yyyy hh:mm"
                else:
                    d = C.to_date(val)
                    if d:
                        tgt.value = d
                        tgt.number_format = "dd/mm/yyyy"
                    else:
                        continue
            else:
                tgt.value = val
            n += 1
    return n


def fill_cers(ctx, ws, res):
    t = ws.title.lower()
    if t not in ("vessel", "voyage", "waste", "security"):
        return False
    n = fill_cers_header(ctx, ws, res)
    if t == "waste":
        import waste2 as W2
        tb = {"row": 22, "cols": {10: "w_delivered_last", 12: "w_deliver", 14: "w_capacity"}, "label": 3, "alt": 3}
        filled, cleared, unmatched = W2.fill_table(ctx, ws, res, tb, keep_unmatched=True)
        res.lines.append(f"{filled} waste lines filled")
    res.kind = f"CERS3 {ws.title}"
    res.lines.append(f"{n} fields filled")
    return True

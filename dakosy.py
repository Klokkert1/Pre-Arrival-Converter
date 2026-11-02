"""DAKOSY (German port reporting workbook): label in column B, input in column C, one sheet per message."""
import re, datetime
from openpyxl.cell.cell import MergedCell
from matching import norm
import convert as C
import health as H

SPEC = [  # (regex on normalised label, key)
    (r"^voyage number of originator", "voyage_no"), (r"^name of reporting party", "master_name"),
    (r"^name of ship agency", "agent"),
    (r"^ship name", "ship_name"), (r"^callsign", "call_sign"), (r"^imo number", "imo"), (r"^mmsi number", "mmsi"),
    (r"^flag", "flag"), (r"^length overall", "length"), (r"^beam", "breadth"), (r"^gross tonnage", "gt"),
    (r"^port of registry", "port_registry"), (r"^inmarsat call number$", "inmarsat"), (r"^ship type", "ship_type"),
    (r"^ism name of company", "company_name"), (r"^ism company identification", "company_imo"),
    (r"^name of master", "master_name"), (r"^position of ship in port of call", "position"),
    (r"^dead weight summer", "dwt"),
    (r"^estimated time of arrival eta port of call", "eta"), (r"^estimated time of departure etd port of call", "etd"),
    (r"^last port lp", "last_port_locode"), (r"^estimated time of departure last port", "etd_last"),
    (r"^next port np", "next_port_locode"),
    (r"^number of persons on board", "persons"), (r"^number of crew members on board", "n_crew"),
    (r"^number of passengers on board", "n_pax"), (r"^number of stowaways", "n_stow"),
    (r"^draught on (arrival|departure) in decimetre", "draught_dm"),
    (r"^current ship security level", "ship_sec_level"), (r"^cso last name", "cso_family"),
    (r"^cso first name", "cso_given"), (r"^cso telephone", "cso_phone"), (r"^cso e ?mail", "cso_email"),
    (r"^valid issc on board", "issc_valid"), (r"^issc type", "issc_type"), (r"^issc issuer type", "issc_issuer_type"),
    (r"^issc issuer name", "issc_issuer"), (r"^date of expiration issc", "issc_expiry"),
    (r"^approved security plan on board", "ssp_onboard"), (r"^port facility of arrival", "port_facility"),
    (r"^general description of cargo", "cargo_desc"),
    (r"^last port where waste", "w_last_locode"), (r"^date of last disposal", "w_last_date"),
    (r"^next waste disposal port", "w_next_locode"),
    (r"^number of deaths", "death_count"), (r"^number of ill persons", "ill_count"),
    (r"^place of issue", "san_place"), (r"^date of issue", "san_date"),
]
YESNO = {"issc_valid", "ssp_onboard"}


def is_dakosy(wb):
    return "main" in wb.sheetnames and "port message" in wb.sheetnames and \
        any("dakosy" in str(wb[n].cell(r, c).value or "").lower()
            for n in ("main", "ship") if n in wb.sheetnames for r in range(1, 12) for c in range(1, 6)) or \
        ("Instructions" in wb.sheetnames and "traffic message on arrival" in wb.sheetnames)


def _vals(m, ctx, ws):
    f = m.fields
    dep = "departure" in ws.title.lower()
    last = m.ports[0] if m.ports else {}
    draft = max(f.get("draft_fore") or 0, f.get("draft_aft") or 0)
    nil = lambda k: 0 if str(f.get(k, "")).strip().lower() in ("nil", "none", "no", "0") else None
    return {
        "voyage_no": f.get("voyage_no"), "master_name": f.get("master_name"), "agent": f.get("agent"),
        "ship_name": f.get("ship_name"), "call_sign": f.get("call_sign"), "imo": f.get("imo"), "mmsi": f.get("mmsi"),
        "flag": f.get("flag"), "length": f.get("length"), "breadth": f.get("breadth"), "gt": f.get("gt"),
        "port_registry": f.get("port_registry"), "inmarsat": f.get("inmarsat"), "ship_type": f.get("ship_type"),
        "company_name": f.get("company_name"), "company_imo": f.get("company_imo"),
        "position": f.get("ship_location") or f.get("berth"), "dwt": f.get("dwt"),
        "eta": f.get("eta_arrival") or f.get("eta"), "etd": f.get("etd_departure"),
        "last_port_locode": f.get("last_port_locode"), "etd_last": last.get("departure"),
        "next_port_locode": f.get("next_port_locode"),
        "persons": (len(m.crew) + len(m.pax)) or None, "n_crew": len(m.crew) or None,
        "n_pax": len(m.pax) if m.crew else None, "n_stow": nil("nil_stowaways"),
        "draught_dm": round(draft * 10) if draft else None,
        "ship_sec_level": f.get("ship_sec_level"), "cso_family": f.get("cso_family"), "cso_given": f.get("cso_given"),
        "cso_phone": f.get("cso_phone"), "cso_email": f.get("cso_email"), "issc_valid": f.get("issc_valid"),
        "issc_type": None, "issc_issuer_type": f.get("issc_issuer_type"), "issc_issuer": f.get("issc_issuer"),
        "issc_expiry": f.get("issc_expiry"), "ssp_onboard": f.get("ssp_onboard"), "port_facility": f.get("port_facility"),
        "cargo_desc": f.get("cargo_desc"),
        "w_last_locode": None, "w_last_date": m.waste["header"].get("w_last_date"),
        "w_next_locode": f.get("next_port_locode"),
        "death_count": m.health.get("death_count"), "ill_count": m.health.get("ill_count"),
        "san_place": f.get("san_place"), "san_date": f.get("san_date"),
    }


def ship_type_code(wb, name):
    if "F1" not in wb.sheetnames or not name:
        return name
    for r in range(1, wb["F1"].max_row + 1):
        if str(wb["F1"].cell(r, 2).value or "").strip().lower() == str(name).strip().lower():
            return str(wb["F1"].cell(r, 1).value).replace(" ", "")
    return name


def yn(v):
    s = str(v).strip().lower()
    if s in ("yes", "y", "true", "ja", "1"):
        return "y"
    if s in ("no", "n", "false", "nee", "0"):
        return "n"
    return None


def fill_dakosy(ctx, ws, res):
    from forms import mark, anchor, span, is_formula, blank
    m = ctx["master"]
    f = m.fields
    vals = _vals(m, ctx, ws)
    if not vals["w_last_locode"]:
        for p in m.ports:
            if p.get("port") and p.get("port") == m.waste["header"].get("w_last_port"):
                vals["w_last_locode"] = p.get("locode")
    n = 0
    t = ws.title.lower()
    for r in range(1, min(ws.max_row, 60) + 1):
        cell = ws.cell(r, 2)
        v = cell.value
        if isinstance(cell, MergedCell) or not isinstance(v, str):
            continue
        lab = norm(v.replace("*", ""))
        c0, c1, r0, r1 = span(ws, r, 2)
        tgt = anchor(ws, r0, c1 + 1)
        if tgt is None or isinstance(tgt, MergedCell) or is_formula(tgt.value):
            continue
        key = next((k for rx, k in SPEC if re.search(rx, lab)), None)
        health_key = None
        if t == "health":
            health_key = next((k for k, rx in H.QUESTIONS if re.search(rx, v.lower())), None)
        if not key and not health_key:
            continue
        if isinstance(tgt.value, str) and any(re.search(rx, norm(tgt.value)) for rx, _ in SPEC):
            continue
        if health_key and not key:
            val = yn(m.health.get(health_key))
            key = "yn_" + health_key
        else:
            val = vals.get(key)
        if key == "flag" and val:
            v2, ok = C.convert_country(val, None, "iso 2 alpha-2")
            val = v2 if ok and v2 else val
        elif key == "ship_type":
            val = ship_type_code(ctx["wb"], val)
        elif key in YESNO:
            val = yn(val) or val
        elif key in ("eta", "etd", "etd_last"):
            if hasattr(val, "hour"):
                pass
            elif val is not None and C.to_date(val):
                val = datetime.datetime.combine(C.to_date(val), datetime.time())
        elif key in ("w_last_date", "issc_expiry", "san_date"):
            d = C.to_date(val)
            val = datetime.datetime.combine(d, datetime.time()) if d else val
        if blank(val):
            if not blank(tgt.value) and not re.fullmatch(r"(TRUE|FALSE|NO|YES|\d+)?", str(tgt.value)):
                tgt.value = None
            if "*" in v and not key.startswith("yn_"):
                mark(tgt); res.yellow += 1; res.missing.append(v.strip(" *")[:40])
            continue
        tgt.value = val
        if isinstance(val, datetime.datetime) and not C.is_date_format(tgt.number_format):
            tgt.number_format = "dd.mm.yyyy hh:mm" if key in ("eta", "etd", "etd_last") else "dd.mm.yyyy"
        n += 1
    res.kind = f"DAKOSY {ws.title}"
    res.lines.append(f"{n} fields filled")
    return n

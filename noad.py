"""EU-NoAD Data Collecting Tool (German NSW): numbered sections, label in B:E, input box to the right."""
import re, datetime
from openpyxl.cell.cell import MergedCell
from matching import norm
import convert as C

# (section prefix or None, regex on normalised label, key)
SPEC = [
    ("1.1", r"^name of last port", "last_port"), ("1.1", r"etd date", "etd_last_d"), ("1.1", r"etd time", "etd_last_t"),
    ("1.3", r"^name of port of call", "port"), ("1.3", r"eta date", "eta_d"), ("1.3", r"eta time", "eta_t"),
    ("1.3", r"etd date", "etd_d"), ("1.3", r"etd time", "etd_t"),
    ("1.5", r"^name of next port", "next_port"),
    ("2.2", r"^position of ship in port", "position"),
    ("2.3", r"^name of agency", "agent"),
    ("2.5", r"^general description of cargo", "cargo_desc"),
    (None, r"^arrival draft|^departure draft", "draft"), (None, r"^current ship security level", "ship_sec_level"),
    ("3.2", r"^total", "persons"), ("3.2", r"^crew", "n_crew"), ("3.2", r"^passenger", "n_pax"), ("3.2", r"^stowaways", "n_stow"),
    ("4.2", r"^total", "persons"), ("4.2", r"^crew", "n_crew"), ("4.2", r"^passenger", "n_pax"), ("4.2", r"^stowaways", "n_stow"),
    ("3.4", r"^dg cargo on board", "dg_cargo"), ("4.4", r"^dg cargo on board", "dg_cargo"),
    ("5.1", r"^name of ship", "ship_name"), ("5.1", r"^imo or eni", "imo"), ("5.1", r"^type of ship", "ship_type"),
    ("5.1", r"^flag", "flag"), ("5.1", r"^port of registry", "port_registry"), ("5.1", r"^call sign", "call_sign"),
    ("5.1", r"^mmsi", "mmsi"), ("5.1", r"^length overall", "length"), ("5.1", r"^beam", "breadth"),
    ("5.1", r"^gross tonnage", "gt"), ("5.1", r"^net tonnage", "nt"), ("5.1", r"^dead weight", "dwt"),
    ("5.1", r"^certificate of registry number", "cert_no"), ("5.1", r"^ism manager.*name of co", "company_name"),
    ("5.1", r"^company ident", "company_imo"), ("5.1", r"^vessel s email", "ship_email"),
    ("5.2", r"^valid sanitation", "yn_valid_cert"), ("5.2", r"^sscec details.*place of issue", "san_place"),
    ("5.2", r"^date of issue", "san_date"), ("5.2", r"^reinspection", "yn_reinspection"),
    ("5.3", r"^cso.*last name", "cso_family"), ("5.3", r"^first name", "cso_given"), ("5.3", r"^telephone", "cso_phone"),
    ("5.3", r"^e mail", "cso_email"), ("5.3", r"^valid issc", "issc_valid"),
    ("5.3", r"^issc issuing details.*type", "issc_type"), ("5.3", r"^expiry date", "issc_expiry"),
    ("5.3", r"^issuer type", "issc_issuer_type"), ("5.3", r"^issued by", "issc_issuer"),
    ("5.3", r"^approved security plan", "ssp_onboard"),
]


def is_noad(wb):
    return "INSTRUCTIONS" in wb.sheetnames and any(n.startswith("1. VOYAGE") for n in wb.sheetnames) and \
        any(n.startswith("5. SHIP") for n in wb.sheetnames)


def yn(v):
    s = str(v).strip().lower()
    return "y" if s in ("yes", "y", "true", "ja", "1") else "n" if s in ("no", "n", "false", "nee", "0") else None


def fill_noad(ctx, ws, res):
    from forms import mark, anchor, span, is_formula, blank
    m = ctx["master"]
    f = m.fields
    last = m.ports[0] if m.ports else {}
    eta = f.get("eta_arrival") or f.get("eta")
    etd = f.get("etd_departure")
    draft = max(f.get("draft_fore") or 0, f.get("draft_aft") or 0) or None
    nil = lambda k: 0 if str(f.get(k, "")).strip().lower() in ("nil", "none", "no", "0") else None
    dtm = lambda v: v if hasattr(v, "hour") else None
    vals = {
        "last_port": f.get("last_port"), "etd_last_d": C.to_date(last.get("departure")),
        "etd_last_t": dtm(last.get("departure")), "port": f.get("port"), "eta_d": C.to_date(eta), "eta_t": dtm(eta),
        "etd_d": C.to_date(etd), "etd_t": dtm(etd), "next_port": f.get("next_port"),
        "position": f.get("ship_location") or f.get("berth"), "agent": f.get("agent"), "cargo_desc": f.get("cargo_desc"),
        "draft": draft, "ship_sec_level": f.get("ship_sec_level"), "persons": (len(m.crew) + len(m.pax)) or None,
        "n_crew": len(m.crew) or None, "n_pax": len(m.pax) if m.crew else None, "n_stow": nil("nil_stowaways"),
        "dg_cargo": yn(f.get("dg_cargo")), "ship_name": f.get("ship_name"), "imo": f.get("imo"),
        "ship_type": f.get("ship_type"), "flag": f.get("flag"), "port_registry": f.get("port_registry"),
        "call_sign": f.get("call_sign"), "mmsi": f.get("mmsi"), "length": f.get("length"), "breadth": f.get("breadth"),
        "gt": f.get("gt"), "nt": f.get("nt"), "dwt": f.get("dwt"), "cert_no": f.get("cert_registry"),
        "company_name": f.get("company_name"), "company_imo": f.get("company_imo"), "ship_email": f.get("ship_email"),
        "yn_valid_cert": ("y" if str(f.get("san_type", "")).lower().startswith(("cert", "exempt")) else None),
        "san_place": f.get("san_place"), "san_date": C.to_date(f.get("san_date")),
        "yn_reinspection": yn(f.get("reinspection")),
        "cso_family": f.get("cso_family"), "cso_given": f.get("cso_given"), "cso_phone": f.get("cso_phone"),
        "cso_email": f.get("cso_email"), "issc_valid": yn(f.get("issc_valid")), "issc_type": None,
        "issc_expiry": C.to_date(f.get("issc_expiry")), "issc_issuer_type": f.get("issc_issuer_type"),
        "issc_issuer": f.get("issc_issuer"), "ssp_onboard": yn(f.get("ssp_onboard")),
    }
    n = 0
    sect = None
    stop = 51 if ws.title.startswith("1.") else 60
    for r in range(1, min(ws.max_row, 90) + 1):
        if r >= stop and ws.title.startswith("1."):
            break
        cell = ws.cell(r, 2)
        v = cell.value
        if isinstance(cell, MergedCell) or not isinstance(v, str):
            continue
        ms = re.match(r"^(\d\.\d)\s", v.strip())
        if ms:
            sect = ms.group(1)
            continue
        lab = norm(re.sub(r"\(.*?\)", " ", v).replace("-", " "))
        key = None
        for sec, rx, k in SPEC:
            if (sec is None or sec == sect) and re.search(rx, lab):
                key = k
                break
        if not key:
            continue
        c0, c1, r0, r1 = span(ws, r, 2)
        tgt = anchor(ws, r0, c1 + 1)
        if tgt is None or isinstance(tgt, MergedCell) or is_formula(tgt.value):
            continue
        val = vals.get(key)
        if key == "flag" and val:
            v2, ok = C.convert_country(val, None, "iso 2 alpha-2")
            val = v2 if ok and v2 else val
        if blank(val):
            if not blank(tgt.value):
                tgt.value = None
            mark(tgt); res.yellow += 1; res.missing.append(re.sub(r"\s+", " ", v).strip()[:40])
            continue
        if key.endswith("_d") or key in ("san_date", "issc_expiry"):
            d = C.to_date(val)
            tgt.value = datetime.datetime.combine(d, datetime.time())
            tgt.number_format = "dd.mm.yyyy"
        elif key.endswith("_t"):
            tgt.value = datetime.time(val.hour, val.minute)
            tgt.number_format = "hh:mm"
        else:
            tgt.value = val
        n += 1
    res.kind = f"EU-NoAD {ws.title}"
    res.lines.append(f"{n} fields filled")
    return n

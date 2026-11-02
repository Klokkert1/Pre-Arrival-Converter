"""MSW (National Single Window) upload templates: header-above-value blocks, parties grid, waste sheet."""
import re, datetime
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import PatternFill
from matching import norm
import convert as C

YELLOW = PatternFill("solid", fgColor="FFFF00")
TECH = re.compile(r"^[A-Z][A-Za-z0-9]*(_[A-Za-z0-9]+)+$")
MSW_SHEETS = {"arrivalinformation", "departureinformation", "arrivalordeparture", "shipdetails", "security",
              "locationsinport"}


def key(title):
    return re.sub(r"[^a-z]", "", title.lower())


def is_msw_book(wb):
    names = {key(n) for n in wb.sheetnames}
    return "version" in names and "crewlist" in names and bool(names & {"shipstores", "wasteandresidues"})


def handles(wb, ws):
    return is_msw_book(wb) and (key(ws.title) in MSW_SHEETS or key(ws.title) in ("wastereceipt", "wasteandresidues"))


def _d(v, fmt="%Y-%m-%d"):
    d = C.to_date(v)
    return d.strftime(fmt) if d else v


# (regex on normalised header, getter(master) -> value)
def LABELS(m):
    f = m.fields
    def g(k): return f.get(k)
    cso = " ".join(str(x) for x in (g("cso_given"), g("cso_family")) if x)
    eta = g("eta_arrival") or g("eta")
    return [
        (r"^(ship ?name|name of ship|name)$", g("ship_name"), "Name"),
        (r"^imo", g("imo"), None),
        (r"^call ?sign", g("call_sign"), None),
        (r"^(flag|flag state)$", g("flag"), None),
        (r"^type\b", g("ship_type"), None),
        (r"^home port|^port of registry", g("port_registry"), None),
        (r"^(gross tonnage|gt)$", g("gt"), None),
        (r"^(net tonnage|nt)$", g("nt"), None),
        (r"^name of master|^master$", g("master_name"), None),
        (r"^(no ?crew|crew)$", len(m.crew) or None, None),
        (r"^(no ?pax|passengers)$", len(m.pax) if m.pax else None, None),
        (r"^(port of call|port)\b.*locode|^port of call$|^port call", g("port_locode") or g("port"), None),
        (r"^(eta|eta port of call|eta from)", _d(eta, "%Y-%m-%d %H:%M"), None),
        (r"^(etd|etd from)", _d(g("etd_departure"), "%Y-%m-%d %H:%M"), None),
        (r"^(purpose of call|primary purpose|purpose)$", g("call_purpose"), None),
        (r"^security level", g("ship_sec_level"), None),
        (r"^(next port|next port of call)", g("next_port_locode") or g("next_port"), None),
        (r"^(last port|last port of call)", g("last_port_locode") or g("last_port"), None),
        (r"^port facility", g("port_facility"), None),
        (r"^(issc number|certificate number)", g("issc_number"), None),
        (r"^(dead ?weight|dwt)", g("dwt"), None), (r"^vessel length|^length", g("length"), None),
        (r"^vessel width|^breadth|^width", g("breadth"), None), (r"^vessel moulded depth|^depth", g("depth"), None),
        (r"^draft fore|^draught fore", g("draft_fore"), None), (r"^draft aft|^draught aft", g("draft_aft"), None),
        (r"^draught level max", (max(g("draft_fore") or 0, g("draft_aft") or 0) or None), None), (r"^engine power", g("engine_power"), None),
        (r"^valid issc", g("issc_valid"), None), (r"^issc issuer$", g("issc_issuer"), None),
        (r"^issc type", g("issc_issuer_type"), None), (r"^expiry date", _d(g("issc_expiry")), None),
        (r"^issc document number", g("issc_number"), None), (r"^issc issuing date", _d(g("issc_issued")), None),
        (r"^ssp onboard", g("ssp_onboard"), None),
        (r"^(cso|sso) family name", g("cso_family"), None), (r"^(cso|sso) given name", g("cso_given"), None),
        (r"^phone 24 7", g("cso_phone"), None), (r"^e ?mail$", g("cso_email"), None),
        (r"^(inf class|ice class)$", None, None), (r"^(mgo|lo|fw)$", None, None),
    ]


PARTY_ROWS = [
    (r"^master", lambda f: {"name": f.get("master_name")}),
    (r"^agent", lambda f: {"name": f.get("agent")}),
    (r"^security officer|^cso|^ship security officer",
     lambda f: {"name": " ".join(str(x) for x in (f.get("cso_given"), f.get("cso_family")) if x) or None,
                "phone": f.get("cso_phone"), "mail": f.get("cso_email"),
                "contact": " ".join(str(x) for x in (f.get("cso_given"), f.get("cso_family")) if x) or None}),
    (r"^(ship )?operating company|^company|^owner|^ship owner", lambda f: {"name": f.get("company_name")}),
]
PARTY_COLS = [(r"^title/?name$|^name$|^title", "name"), (r"^telephone|^phone", "phone"),
              (r"^e-?mail", "mail"), (r"^contact person", "contact")]


def _is_label(v):
    return isinstance(v, str) and v.strip() and not v.strip().startswith("=")


def _put(ws, r, c, value, res, stats):
    cell = ws.cell(r, c)
    if isinstance(cell, MergedCell):
        return
    old = cell.value
    if value in (None, ""):
        if old not in (None, "") and not (isinstance(old, str) and old.startswith("=")):
            cell.value = None
            cell.fill = YELLOW
            stats["missing"] += 1
        return
    cell.value = value
    if isinstance(value, (int, float)) is False and re.match(r"^\d{4}-\d\d-\d\d", str(value)):
        cell.number_format = "@"
    stats["filled"] += 1


def fill_msw(ctx, ws, res):
    m = ctx["master"]
    stats = {"filled": 0, "missing": 0}
    labels = LABELS(m)
    f = m.fields
    maxr = min(ws.max_row, 60)
    maxc = min(ws.max_column, 14)
    cells = [(r, c, ws.cell(r, c).value) for r in range(1, maxr + 1) for c in range(1, maxc + 1)]
    cells = [(r, c, v) for r, c, v in cells if isinstance(v, str)]
    # technical-header sheets (Denmark): technical row, readable row, then the data row
    tech_rows = {r for r, c, v in cells if TECH.match(v.strip())}
    tech_rows = {r for r in tech_rows if sum(1 for rr, cc, vv in cells if rr == r and TECH.match(vv.strip())) >= 3}
    done = set()
    # parties grid
    for r, c, v in cells:
        if re.match(r"^title/?name$", norm(v)) or norm(v) == "title name":
            cols = {}
            for rr, cc, vv in cells:
                if rr == r:
                    for rx, k in PARTY_COLS:
                        if re.search(rx, norm(vv)) and k not in cols:
                            cols[k] = cc
            for r2 in range(r + 1, r + 8):
                lab = ws.cell(r2, 1).value
                if not _is_label(lab):
                    continue
                for rx, fn in PARTY_ROWS:
                    if re.search(rx, norm(lab)):
                        vals = fn(f)
                        done.add((r2, 1))
                        allc = {cc for rr, cc, vv in cells if rr == r and cc > 1}
                        for cc in sorted(allc):
                            k = next((kk for kk, c3 in cols.items() if c3 == cc), None)
                            _put(ws, r2, cc, vals.get(k) if k else None, res, stats)
                            done.add((r2, cc))
                        break
            done.add((r, c))
    barge_rows = {r + 1 for r, c, v in cells if c == 1 and norm(v) == "barge"} | \
                 {r + 2 for r, c, v in cells if c == 1 and norm(v) == "barge"}
    for r, c, v in cells:
        if (r, c) in done or r in barge_rows:
            continue
        txt = v.strip()
        nt = norm(txt.replace("_", " ").replace("*", ""))
        tech = r in tech_rows
        if r in tech_rows and not TECH.match(txt):
            continue                       # readable row under the technical row: leave alone
        for rx, val, _ in labels:
            if re.search(rx, nt):
                if r - 1 in tech_rows or (r in tech_rows and False):
                    break
                vr = r + (2 if r in tech_rows else 1)
                below = ws.cell(vr, c).value
                if isinstance(below, str) and not below.startswith("=") and re.search(
                        r"^(name|type|flag|gross|net|dead|home|draft|vessel)\b", norm(below)) and vr - r == 1 and below.endswith(("*", ":")):
                    break
                if r in tech_rows and key(txt) in ("ad",):
                    val = "D" if str(f.get("movement", "")).lower().startswith("dep") else "A"
                if r in tech_rows and "ata" in nt.split() or r in tech_rows and "atd" in nt.split():
                    break
                if not tech and txt.rstrip("* ").endswith(":"):
                    vr, vc = r, c + 1
                    cur = ws.cell(vr, vc).value
                    if re.match(r"^security level", norm(str(cur or ""))) and val not in (None, ""):
                        val = f"Security level {val}"
                else:
                    vc = c
                _put(ws, vr, vc, val, res, stats)
                done.add((vr, vc))
                break
    res.kind = "MSW template"
    res.lines.append(f"{stats['filled']} fields filled" + (f", {stats['missing']} cleared (not in master)" if stats["missing"] else ""))
    return stats["filled"] > 0

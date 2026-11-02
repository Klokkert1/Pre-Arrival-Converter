"""Maritime Declaration of Health (MDoH): header fields, Yes/No questions, ports of call."""
import re, zipfile, posixpath
import xml.etree.ElementTree as ET
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import PatternFill

import convert as C
from matching import norm

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "p": "http://schemas.openxmlformats.org/package/2006/relationships"}
RID = "{%s}id" % NS["r"]


def read_checkboxes(path, sheet_title):
    """Tick boxes (form controls / ActiveX) of one sheet: [{'link','col','row','width_cols'}] (1-based col/row)."""
    out = []
    try:
        z = zipfile.ZipFile(path)
        wbx = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rid = next((s.get(RID) for s in wbx.find("m:sheets", NS) if s.get("name") == sheet_title), None)
        target = next((r.get("Target") for r in rels if r.get("Id") == rid), None)
        if not target:
            return out
        sheet_file = posixpath.normpath(posixpath.join("xl", target)) if not target.startswith("/") else target[1:]
        rel_file = posixpath.join(posixpath.dirname(sheet_file), "_rels", posixpath.basename(sheet_file) + ".rels")
        if rel_file not in z.namelist():
            return out
        for r in ET.fromstring(z.read(rel_file)):
            if r.get("Type", "").endswith("/vmlDrawing"):
                vml = posixpath.normpath(posixpath.join(posixpath.dirname(sheet_file), r.get("Target")))
                text = z.read(vml).decode("utf-8", "ignore")
                for shape in re.findall(r"<x:ClientData.*?</x:ClientData>", text, re.S):
                    link = re.search(r"<x:FmlaLink>\s*([$A-Z]+\$?\d+)\s*</x:FmlaLink>", shape)
                    anc = re.search(r"<x:Anchor>\s*([\d,\s-]+)</x:Anchor>", shape)
                    if link and anc:
                        n = [int(x) for x in anc.group(1).replace("\n", " ").split(",")]
                        out.append({"link": link.group(1).replace("$", ""), "col": n[0] + 1, "col2": n[4] + 1,
                                    "row": (n[2] + n[6]) // 2 + 1 if n[2] != n[6] else n[2] + 1, "row0": n[2] + 1, "row1": n[6] + 1})
    except Exception:
        return []
    return out


YELLOW = PatternFill("solid", fgColor="FFFF00")

# order matters: first match wins
QUESTIONS = [
    ("rats", r"rats or mice|abnormal mortality"),
    ("valid_cert", r"valid sanitation|sanitation control (exemption|certificate)|derat"),
    ("reinspection", r"re-?inspection"),
    ("affected", r"affected area"),
    ("ill_pax", r"ill passengers|greater than normal"),
    ("death", r"died|death"),
    ("ill_now", r"(ill|sick) person on board now|any (ill|sick) person"),
    ("doctor", r"medical pa?r?actitioner|practitioner|consulted"),
    ("condition", r"condition on board|lead to infection"),
    ("sanitary", r"sanitary measure"),
    ("stowaways", r"stowaways? (been )?found|have any stowaways"),
    ("animal", r"sick animal|pet on board"),
    ("infectious", r"infectious|plague, cholera|suspect"),
    ("doctor_onboard", r"doctor on board|surgeon on board"),
]
DETAILS = [
    ("death_count", r"total no\.? ?of deaths|number of deaths"),
    ("ill_count", r"how many ill"),
    ("san_measure", r"specify type, place and date"),
    ("stow_where", r"where did they join|where did.*join the ship"),
    ("affected_where", r"port and date of visit"),
]
FIELDS = [
    ("port_registry", r"port of registry"),
    ("gt_nt", r"^gt ?/ ?nt"),
    ("gt", r"^gross tonnage \(ship\)|^gross tonnage|^gt\b"),
    ("imo", r"\bimo\b"),
    ("ship_name", r"^(name of ship|m/v|mv|vessel name|name of vessel)"),
    ("flag", r"^\(?nationality|flag"),
    ("last_port", r"arriving from|coming port|coming from"),
    ("next_port", r"sailing to|^next port"),
    ("master_name", r"master.?s name|name of master|^signed$"),
    ("crew_count", r"number of crew|crew members on board|crew on board"),
    ("pax_count", r"number of passengers|passengers on board"),
    ("san_place", r"^issued at|^place$"),
    ("sail_date", r"sailing date"),
    ("eta", r"^eta\b"),
    ("port", r"submitted at the port|^port of arrival"),
]
PORT_LIST = r"list (of )?ports of call|ports of call from commencement|list of ports called"


def yes_no(v):
    if v is None or (isinstance(v, str) and not v.strip()):
        return None
    t = str(v).strip().lower()
    if t[:1] in ("y", "j", "s") or t in ("true", "1", "x"):
        return True
    if t[:1] == "n" or t in ("false", "0"):
        return False
    return None


def is_health(ws):
    text = " ".join(str(ws.cell(r, c).value) for r in range(1, 8) for c in range(1, 8) if ws.cell(r, c).value).lower()
    return bool(re.search(r"declaration of health|health declaration|maritime health", text)) or \
        bool(re.search(r"declaration of health|health declaration", ws.title.lower()))


def _label_of(v):
    t = re.sub(r"\s+", " ", str(v)).strip().lower().rstrip(":").strip()
    return t


LABEL_RX = re.compile("|".join([rx for _, rx in QUESTIONS + DETAILS + FIELDS] + [r"^date$", r"^(place|luogo|data|port)$", PORT_LIST]))


def fill_health(ctx, ws, res, exclude=()):
    from forms import mark, blank, anchor, span, has_border, has_dv

    def vcell(r, c, below=False):
        """(cell, ok) - the cell that takes the value of the label at (r, c); old values are overwritten."""
        c0, c1, r0, r1 = span(ws, r, c)
        if below:
            rr = r1 + 1
            a = anchor(ws, rr, c0)
            for _ in range(3):
                if isinstance(a.value, str) and a.value.strip().endswith(":"):
                    rr += 1
                    a = anchor(ws, rr, c0)
            return a, True
        first_blank = None
        col = c1 + 1
        while col <= min(c1 + 24, ws.max_column):
            a = anchor(ws, r0, col)
            ac0, ac1, _, _ = span(ws, a.row, a.column)
            v = a.value
            if isinstance(v, str) and v.strip() and LABEL_RX.search(_label_of(v)):
                break
            if not blank(v) or has_border(a) or (not a.protection.locked) or has_dv(ws, a):
                return a, True
            if first_blank is None:
                first_blank = a
            col = ac1 + 1
        return first_blank, first_blank is not None

    def find_value_cell(ws_, r, c, force_below=False):
        cell, ok = vcell(r, c, below=force_below)
        return cell, ok, False
    master = ctx["master"]
    h = master.health
    f = master.fields
    boxes = read_checkboxes(ctx.get("path"), ws.title) if ctx.get("path") else []
    # pairs of tick boxes: left = Yes, right = No
    pairs = []
    used = set()
    for b in sorted(boxes, key=lambda x: (x["row0"], x["col"])):
        if id(b) in used:
            continue
        mate = next((o for o in boxes if o is not b and id(o) not in used and abs(o["row0"] - b["row0"]) <= 1
                     and o["col"] > b["col"] and o["col"] - b["col"] <= 5), None)
        if mate:
            used.update({id(b), id(mate)})
            pairs.append((b, mate))
    pairs.sort(key=lambda p: p[0]["row0"])
    taken = set()

    def pair_for(row):
        best = None
        for i, (a, b) in enumerate(pairs):
            if i in taken:
                continue
            if a["row0"] - 1 <= row <= a["row1"]:
                d = abs(a["row0"] - row)
                if best is None or d < best[0]:
                    best = (d, i)
        return None if best is None else best[1]

    ans = {"valid_cert": True if f.get("san_type") else None, "reinspection": yes_no(f.get("reinspection"))}
    for k in ("death", "infectious", "ill_pax", "ill_now", "doctor", "condition", "sanitary", "stowaways", "animal",
              "affected", "rats", "doctor_onboard"):
        ans[k] = yes_no(h.get(k))
    stats = {"fields": 0, "q": 0, "q_total": 0}
    ports_done = False

    def put(cell, value, label):
        if label in ("date", "eta", "sail_date", "san_date"):
            d = C.to_date(value)
            if d:
                if not C.is_date_format(cell.number_format):
                    cell.number_format = "dd-mm-yyyy"
                cell.value = d
                return
        cell.value = value

    def tick(i, answer, label, qcell=None):
        a, b = pairs[i]
        taken.add(i)
        targets = []
        for box in (a, b):
            cell = ws.cell(box["row"], box["col"])
            if isinstance(cell, MergedCell):
                for m in ws.merged_cells.ranges:
                    if cell.coordinate in m:
                        cell = ws.cell(m.min_row, m.min_col)
                        break
            targets.append(cell)
        if any(not blank(t.value) for t in targets) and qcell is not None:
            txt = {True: "YES", False: "NO", None: "[ ] Yes   [ ] No"}[answer]
            qcell.value = f"{qcell.value}   >> {txt}"
            if answer is None:
                mark(qcell); res.yellow += 1; res.missing.append(label)
            else:
                stats["q"] += 1
            return
        for box, word, yes in ((a, "Yes", True), (b, "No", False)):
            cell = ws.cell(box["row"], box["col"])
            if isinstance(cell, MergedCell):
                for m in ws.merged_cells.ranges:
                    if cell.coordinate in m:
                        cell = ws.cell(m.min_row, m.min_col)
                        break
            if not blank(cell.value):
                res.note(f"Could not write the {word} answer for '{label}' (cell not empty).")
                continue
            mark_it = (answer is None)
            cell.value = f"[{'X' if answer is yes else ' '}] {word}"
            if mark_it:
                mark(cell)
        if answer is None:
            res.yellow += 2
            res.missing.append(label)
        else:
            stats["q"] += 1

    master_name = f.get("master_name")
    mi = {"port": f.get("port"), "ship_name": f.get("ship_name"), "imo": f.get("imo"), "flag": f.get("flag"),
          "gt": f.get("gt"), "port_registry": f.get("port_registry"), "last_port": f.get("last_port"),
          "next_port": f.get("next_port"), "master_name": master_name, "san_place": f.get("san_place"),
          "eta": f.get("eta_arrival") or f.get("eta"),
          "sail_date": (master.ports[0].get("departure") if master.ports else None),
          "crew_count": len(master.crew) if master.crew else None,
          "pax_count": len(master.pax) if master.pax else 0,
          "gt_nt": (f"{f.get('gt')} GT / {f.get('nt')} NT" if f.get("gt") and f.get("nt") else f.get("gt"))}
    has_ports_table = bool(ctx.get("has_ports_table"))
    seen_cells = set()
    for r in range(1, ws.max_row + 1):
        for c in range(1, min(ws.max_column, 40) + 1):
            if (r, c) in exclude or (r, c) in seen_cells:
                continue
            cell = ws.cell(r, c)
            v = cell.value
            if isinstance(cell, MergedCell) or not isinstance(v, str) or not v.strip():
                continue
            raw = re.sub(r"\s+", " ", v).strip()
            lab = _label_of(v)
            if "…" in raw or "..." in raw or len(raw) > 200:
                continue
            if re.match(r"^(clear health condition|deseases|diseases)", lab):
                res.note("Free-text health sections (clear health condition, diseases/deaths) are not filled - please complete by hand.")
            # --- ports of call as one text cell
            if re.search(PORT_LIST, lab) and not ports_done:
                ports_done = True
                if has_ports_table:
                    continue
                tgt, looks, filled = find_value_cell(ws, r, c, force_below=True)
                if tgt is not None and not filled and master.ports:
                    txt = "; ".join(f"{i + 1}. {C.to_date(p.get('departure') or p.get('arrival')).strftime('%Y-%m-%d') if C.to_date(p.get('departure') or p.get('arrival')) else ''} "
                                    f"{p.get('locode') or p.get('port')}".replace("  ", " ")
                                    for i, p in enumerate(master.ports[:10]))
                    tgt.value = txt
                    stats["fields"] += 1
                elif tgt is not None and not filled:
                    mark(tgt); res.yellow += 1; res.missing.append(raw[:40])
                continue
            # --- details of a Yes answer
            dcon = next((k for k, rx in DETAILS if re.search(rx, lab)), None)
            if dcon:
                tgt, looks, filled = find_value_cell(ws, r, c)
                if tgt is None or filled:
                    continue
                if h.get(dcon) not in (None, ""):
                    tgt.value = h.get(dcon); stats["fields"] += 1
                else:
                    asked = {"death_count": "death", "ill_count": "ill_pax", "san_measure": "sanitary",
                             "stow_where": "stowaways", "affected_where": "affected"}[dcon]
                    if ans.get(asked) is not False:
                        mark(tgt); res.yellow += 1
                        if ans.get(asked):
                            res.missing.append(raw[:50])
                continue
            # --- Yes/No questions
            qcon = None
            if len(raw) >= 14 and not lab.startswith("if yes"):
                qcon = next((k for k, rx in QUESTIONS if re.search(rx, lab)), None)
            if qcon and not re.match(r"^(note|in the absence)", lab):
                stats["q_total"] += 1
                i = pair_for(r)
                if i is not None:
                    tick(i, ans.get(qcon), raw[:60], cell)
                else:
                    tgt, looks, filled = find_value_cell(ws, r, c)
                    if tgt is not None and not filled and (looks or "yes" in lab):
                        a = ans.get(qcon)
                        if a is None:
                            mark(tgt); res.yellow += 1; res.missing.append(raw[:50])
                        else:
                            tgt.value = "Yes" if a else "No"; stats["q"] += 1
                    else:
                        stats["q_total"] -= 1
                continue
            # --- header fields
            key = None
            if lab == "date":
                near = any("issued at" in _label_of(ws.cell(r, cc).value or "") for cc in range(1, 40))
                ctxt = " ".join(_label_of(ws.cell(rr, cc).value or "") for rr in (r - 1, r) for cc in range(1, 6))
                near = near or bool(re.search(r"derat|sanit|exempt", ctxt))
                key = "san_date" if near else "eta"
                if r > 30:
                    key = "eta"
            else:
                key = next((k for k, rx in FIELDS if re.search(rx, lab)), None)
            if key == "doctor_onboard":
                key = None
            if re.search(r"doctor on board", lab):
                tgt, looks, filled = find_value_cell(ws, r, c)
                if tgt is not None and not filled:
                    a = ans.get("doctor_onboard")
                    if a is None:
                        mark(tgt); res.yellow += 1; res.missing.append("Doctor on board")
                    else:
                        tgt.value = "Yes" if a else "No"; stats["fields"] += 1
                continue
            if not key:
                continue
            if key == "san_date":
                val = f.get("san_date")
            elif key == "eta" and re.search(r"hour|time|ora", lab):
                val = mi.get("eta")
                tgt0, _, _ = find_value_cell(ws, r, c)
                if tgt0 is not None and val is not None and hasattr(val, "hour"):
                    tgt0.number_format = "dd-mm-yyyy hh:mm"
                    tgt0.value = val; stats["fields"] += 1; seen_cells.add((tgt0.row, tgt0.column))
                    continue
            else:
                val = mi.get(key)
            tgt, looks, filled = find_value_cell(ws, r, c)
            if key in ("san_date", "san_place") and re.search(r"derat|sanit|exempt", " ".join(
                    _label_of(ws.cell(rr, cc).value or "") for rr in (r - 1, r) for cc in range(1, 6))):
                tgt, looks, filled = find_value_cell(ws, r, c, force_below=True)
            if tgt is None or filled:
                continue
            if blank(val):
                mark(tgt); res.yellow += 1; res.missing.append(raw[:50])
                continue
            put(tgt, val, "date" if key in ("san_date", "eta", "sail_date") else key)
            seen_cells.add((tgt.row, tgt.column))
            stats["fields"] += 1
    unmatched = [i for i in range(len(pairs)) if i not in taken]
    if unmatched:
        res.note(f"{len(unmatched)} tick-box pair(s) of this form could not be linked to a question - please check them by hand.")
    if boxes:
        res.note("The tick boxes of this form cannot be kept in the copy: the answers are written as text "
                 "([X] Yes / [ ] No) at the same place. Please check them.")
    res.kind = "Maritime Declaration of Health"
    res.lines.append(f"{stats['fields']} fields and {stats['q']} of {stats['q_total']} questions filled")
    return stats

"""Recognising and filling port forms (Excel)."""
import re, copy, zipfile, datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter as L

from matching import match, norm, compact, hint_text
import vocabulary as V
import convert as C
import health as H
import stores as ST
import waste2 as W2

YELLOW = PatternFill("solid", fgColor="FFFF00")
DATE_CONCEPTS = {"dob", "id_issue_date", "id_expiry", "date_emb", "date_disemb", "arrival", "departure"}
COUNTRY_CONCEPTS = {"nationality", "id_issue_country", "country"}


class SheetResult:
    def __init__(self, name):
        self.name, self.kind = name, "not recognised"
        self.lines, self.yellow, self.unrecognised, self.missing, self.notes = [], 0, [], [], []

    def note(self, s):
        if s not in self.notes:
            self.notes.append(s)


# ------------------------------------------------------------------ small helpers
def has_border(cell):
    b = cell.border
    return any(getattr(b, side) is not None and getattr(b, side).style for side in ("left", "right", "top", "bottom"))

def span(ws, r, c):
    for m in ws.merged_cells.ranges:
        if m.min_row <= r <= m.max_row and m.min_col <= c <= m.max_col:
            return m.min_col, m.max_col, m.min_row, m.max_row
    return c, c, r, r

def anchor(ws, r, c):
    c0, c1, r0, r1 = span(ws, r, c)
    return ws.cell(r0, c0)

def blank(v):
    return v is None or (isinstance(v, str) and not v.strip())

def is_formula(v):
    return isinstance(v, str) and v.startswith("=")

def mark(cell):
    cell.fill = YELLOW

def get_allowed(wb, ws, cell):
    """Values of the dropdown list on a cell, or None."""
    for dv in ws.data_validations.dataValidation:
        if dv.type == "list" and cell.coordinate in dv.sqref and dv.formula1:
            f = dv.formula1.strip()
            if f.startswith('"'):
                return [x.strip() for x in f.strip('"').split(",")]
            f = f.lstrip("=")
            if f in wb.defined_names:
                f = wb.defined_names[f].attr_text
            if "[" in f or "#REF" in f:
                return None
            if "!" in f:
                sh, rng = f.rsplit("!", 1)
                sh = sh.strip("'")
                if sh not in wb.sheetnames:
                    return None
                src = wb[sh]
            else:
                src, rng = ws, f
            rng = rng.replace("$", "")
            try:
                cells = src[rng]
            except Exception:
                return None
            if not isinstance(cells, tuple):
                return [cells.value]
            vals = []
            for row in cells:
                for c in (row if isinstance(row, tuple) else (row,)):
                    if c.value not in (None, ""):
                        vals.append(c.value)
            return vals
    return None

def has_dv(ws, cell):
    return any(cell.coordinate in dv.sqref for dv in ws.data_validations.dataValidation)

def package_notes(path):
    notes = []
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            if any(b"x14:dataValidation" in z.read(n) for n in names if n.startswith("xl/worksheets/sheet")):
                notes.append("Some dropdown lists in this form use an Excel extension format that cannot be kept in the "
                             "copy. The values are filled in, but those cells lose their dropdown arrow.")
            if any(n.startswith("xl/ctrlProps") for n in names):
                notes.append("This form contains form controls (checkboxes/buttons) that are not kept in the copy.")
    except Exception:
        pass
    return notes

def parse_instructions(wb):
    """Read the field rules from an Instructions sheet (as in the NSW forms): label -> rule text."""
    rules = {}
    for ws in wb.worksheets:
        if "instruction" not in ws.title.lower():
            continue
        rows = []
        for r in range(1, min(ws.max_row, 400) + 1):
            rows.append([str(c.value).strip() for c in ws[r][:8] if not blank(c.value)])
        label, buf = None, []
        for i, texts in enumerate(rows):
            if not texts:
                continue
            nxt = rows[i + 1][0] if i + 1 < len(rows) and rows[i + 1] else ""
            if nxt.lower().startswith("field status") and not texts[0].lower().startswith("field"):
                if label:
                    rules[norm(label)] = " ".join(buf)
                label, buf = texts[0], []
            elif label:
                buf.append(" ".join(texts))
        if label:
            rules[norm(label)] = " ".join(buf)
    return rules

def rule_for(ctx, texts):
    from matching import variants
    for x in texts:
        for v in variants(x):
            r = ctx["rules"].get(v)
            if r:
                return r
    return ""


# ------------------------------------------------------------------ table detection
class Table:
    pass

def detect_tables(ws, vocab, min_hits=3):
    hits = {}
    for r in range(1, min(ws.max_row, 40) + 1):
        m = {}
        for c in range(1, min(ws.max_column, 40) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and 0 < len(v.strip()) < 90:
                concept, s = match(v, vocab)
                if concept and len(v.strip()) <= 45:
                    m[c] = (concept, s)
        concepts = {x[0] for x in m.values()}
        if len(concepts) >= min_hits and len(concepts - {"serial"}) >= 3:
            hits[r] = m
    groups = []
    for r in sorted(hits):
        if groups and r - groups[-1][-1] <= 1:
            groups[-1].append(r)
        else:
            groups.append([r])
    tables = []
    for g in groups:
        t = Table()
        t.rows = g
        t.cols, t.hints, t.unmatched, t.mandatory = {}, {}, [], set()
        best = {}
        texts = {}
        for r in g:
            for c in range(1, min(ws.max_column, 40) + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str) and v.strip():
                    texts.setdefault(c, []).append(v)
                    concept, s = match(v, vocab)
                    if concept and s > best.get(c, (None, 0))[1]:
                        best[c] = (concept, s)
        used = set()
        for c, (concept, s) in sorted(best.items(), key=lambda kv: -kv[1][1]):
            if concept not in used or concept in ("serial",):
                t.cols[c] = concept
                used.add(concept)
        t.texts = texts
        for c, tx in texts.items():
            t.hints[c] = " ".join(tx)
            if any("*" in x for x in tx):
                t.mandatory.add(c)
            if c not in t.cols and any(len(x.strip()) > 1 for x in tx):
                if len(" ".join(tx)) < 80:
                    t.unmatched.append(c)
        data_start = g[-1] + 1
        keycols = list(t.cols)
        while any(isinstance(ws.cell(data_start, c), MergedCell) for c in keycols[:3]):
            data_start += 1
        t.data_start = data_start
        t.serial_col = next((c for c, k in t.cols.items() if k == "serial"), None)
        sc = t.serial_col or (min(keycols) - 1 if keycols and min(keycols) > 1 else None)
        if sc:
            # a second header line in another language: text only, while the lines below are numbered
            for _ in range(2):
                nxt = ws.cell(t.data_start + 1, sc).value
                texts_only = [ws.cell(t.data_start, c).value for c in keycols
                              if not isinstance(ws.cell(t.data_start, c), MergedCell)]
                texts_only = [x for x in texts_only if not blank(x)]
                if blank(ws.cell(t.data_start, sc).value) and not blank(nxt) and re.fullmatch(r"\d{1,3}", str(nxt).strip()) \
                        and texts_only and all(isinstance(x, str) and not re.search(r"\d", x) for x in texts_only):
                    t.data_start += 1
        r, cap = t.data_start, 0
        while r <= ws.max_row + 1:
            row_has = any(has_border(ws.cell(r, c)) for c in keycols[:3])
            if t.serial_col and not blank(ws.cell(r, t.serial_col).value):
                row_has = True
            if not row_has:
                break
            cap += 1
            r += 1
        t.capacity = cap
        tables.append(t)
    return tables

_REF = re.compile(r"((?:'[^']+'|[A-Za-z0-9_]+)!)?(?<![A-Za-z0-9_])(\$?[A-Z]{1,3}\$?)(\d+)(?::(\$?[A-Z]{1,3}\$?)(\d+))?(?![A-Za-z0-9_(])")

def adjust_formula(formula, after_row, n):
    """After inserting n rows below `after_row`: move references to rows below it, and let a range that
    ends on `after_row` grow with the table."""
    def fix(m):
        if m.group(1):          # reference to another sheet: leave alone
            return m.group(0)
        c1, r1, c2, r2 = m.group(2), int(m.group(3)), m.group(4), m.group(5)
        if c2 is None:
            return f"{c1}{r1 + n if r1 > after_row else r1}"
        r2 = int(r2)
        new1 = r1 + n if r1 > after_row else r1
        new2 = r2 + n if r2 >= after_row else r2
        return f"{c1}{new1}:{c2}{new2}"
    return _REF.sub(fix, formula)

def shift_rows_down(ws, after_row, n):
    """Make room: move everything below `after_row` down by n rows (merged cells and row heights follow)."""
    merged = [m for m in list(ws.merged_cells.ranges) if m.min_row > after_row]
    for m in merged:
        ws.unmerge_cells(str(m))
    last_col = ws.max_column
    ws.move_range(f"A{after_row + 1}:{L(last_col)}{ws.max_row}", rows=n, translate=False)
    for row in ws.iter_rows():
        for cell in row:
            if is_formula(cell.value):
                cell.value = adjust_formula(cell.value, after_row, n)
    for m in merged:
        ws.merge_cells(start_row=m.min_row + n, start_column=m.min_col, end_row=m.max_row + n, end_column=m.max_col)
    for r in range(ws.max_row, after_row, -1):
        h = ws.row_dimensions[r].height if r in ws.row_dimensions else None
        if h is not None:
            ws.row_dimensions[r + n].height = h
    for r in range(after_row + 1, after_row + n + 1):
        ws.row_dimensions[r].height = None

def extend_table(ws, t, need, res):
    cap = t.capacity
    last = t.data_start + cap - 1 if cap > 0 else t.data_start - 1
    tmpl = last if cap > 0 else t.data_start
    below = any(not blank(ws.cell(r, c).value) for r in range(last + 1, ws.max_row + 1)
                for c in range(1, ws.max_column + 1))
    # copy the template row first (cells below may be shifted)
    tmpl_cells = [(c, copy.copy(ws.cell(tmpl, c)._style)) for c in range(1, ws.max_column + 1)]
    tmpl_h = ws.row_dimensions[tmpl].height if tmpl in ws.row_dimensions else None
    tmpl_merges = [m for m in ws.merged_cells.ranges if m.min_row == m.max_row == tmpl]
    if below:
        shift_rows_down(ws, last, need)
    for k in range(1, need + 1):
        r = last + k
        for c, st in tmpl_cells:
            cell = ws.cell(r, c)
            if not isinstance(cell, MergedCell):
                cell._style = copy.copy(st)
        ws.row_dimensions[r].height = tmpl_h
        for m in tmpl_merges:
            ws.merge_cells(start_row=r, start_column=m.min_col, end_row=r, end_column=m.max_col)
    res.note(f"Added {need} extra row(s) below the table (formatting copied from the row above)"
             f"{'; content below was moved down' if below else ''}. Dropdown lists on the new rows are not extended.")
    t.capacity = cap + need


# ------------------------------------------------------------------ value conversion for table cells
RANK_NORM = {C.norm(k): v for k, v in C.RANK_ALIASES.items()}

def convert_cell(ctx, ws, cell, concept, rec, hints, mandatory, rule):
    """Return (value, status). status: ok | blank | missing | unsure."""
    if concept in V.NOT_IN_MASTER:
        return None, "missing"
    if concept == "port":
        want_locode = "locode" in norm(hints)
        raw = (rec.get("locode") or rec.get("port")) if want_locode else rec.get("port")
        return (raw, "ok") if not blank(raw) else (None, "missing" if mandatory else "blank")
    raw = rec.get(concept)
    if blank(raw):
        return None, ("missing" if mandatory else "blank")
    wb = ctx["wb"]
    if concept in DATE_CONCEPTS or (concept in ("w_last_date",)):
        d = C.to_date(raw)
        if d is None:
            return raw, "unsure"
        fmt = C.date_hint_format(hints)
        if fmt:
            cell.number_format = fmt
        elif not C.is_date_format(cell.number_format):
            cell.number_format = "dd-mm-yyyy"
        return d, "ok"
    allowed = get_allowed(wb, ws, cell)
    if concept in COUNTRY_CONCEPTS:
        extra = " iso 3" if "alpha-3" in rule.lower() else (" iso 2 code" if "alpha-2" in rule.lower() else "")
        v, ok = C.convert_country(raw, allowed, hints + extra)
        return (v, "ok") if ok else (raw, "unsure")
    if concept == "rank":
        if allowed:
            pick = C.choose_from_list(raw, allowed, RANK_NORM)
            return (pick, "ok") if pick else (raw, "unsure")
        return raw, "ok"
    if concept == "id_type":
        codes = re.findall(r"\b([A-Z])\s*\(([^)]*)\)", rule)
        if codes and not allowed:
            t = norm(raw)
            for code, name in codes:
                if ("passport" in norm(name) and "passport" in t) or ("id card" in norm(name) and "id card" in t):
                    return code, "ok"
            return raw, "unsure"
        if allowed:
            pick = C.choose_from_list(raw, allowed)
            return (pick, "ok") if pick else (raw, "unsure")
        return raw, "ok"
    if concept == "gender":
        v, ok = C.convert_gender(raw, hints, allowed)
        return (v, "ok") if ok else (raw, "unsure")
    if concept == "transit":
        v, ok = C.convert_yesno(raw, hints, allowed)
        return (v, "ok") if ok else (raw, "unsure")
    if concept in ("family_name", "given_names"):
        if "capital letters" in rule.lower():
            return str(raw).upper(), "ok"
        return raw, "ok"
    if concept in ("sec_level",):
        return (int(raw) if str(raw).isdigit() else raw), "ok"
    return raw, "ok"


def fill_table(ctx, ws, t, records, res):
    if not records:
        return 0
    # skip tables that are driven by formulas
    probe = [ws.cell(t.data_start, c).value for c in t.cols]
    if any(is_formula(v) for v in probe):
        res.note("This sheet is filled by formulas from another sheet; left untouched.")
        return 0
    if len(records) > t.capacity:
        extend_table(ws, t, len(records) - t.capacity, res)
    yellow_missing = set()
    for i, rec in enumerate(records):
        row = t.data_start + i
        for c, concept in t.cols.items():
            cell = ws.cell(row, c)
            if isinstance(cell, MergedCell):
                cell = anchor(ws, row, c)
            if concept == "serial":
                if blank(cell.value):
                    cell.value = i + 1
                continue
            rule = rule_for(ctx, t.texts.get(c, []))
            mand = c in t.mandatory or bool(re.search(r"is mandatory\.", rule, re.I))
            val, status = convert_cell(ctx, ws, cell, concept, rec, t.hints.get(c, ""), mand, rule)
            if status == "ok" or (status == "unsure" and val is not None):
                if blank(cell.value) or not is_formula(cell.value):
                    cell.value = val
            if status in ("missing", "unsure"):
                mark(cell)
                res.yellow += 1
                if status == "missing":
                    yellow_missing.add(t.texts[c][-1].strip().lstrip("*# ").strip() if c in t.texts else concept)
                else:
                    res.note(f"Value not recognised for '{t.texts[c][-1].strip().lstrip('*# ').strip()}' "
                             f"(master value copied, please check).")
        for c in t.unmatched:
            cell = ws.cell(row, c)
            if isinstance(cell, MergedCell):
                cell = anchor(ws, row, c)
            mark(cell)
            res.yellow += 1
    cleared = 0
    for row in range(t.data_start + len(records), t.data_start + t.capacity):
        for c, concept in t.cols.items():
            if concept == "serial":
                continue
            cell = ws.cell(row, c)
            if isinstance(cell, MergedCell):
                continue
            if not blank(cell.value) and not is_formula(cell.value):
                cell.value = None
                cleared += 1
    if cleared:
        res.note(f"{cleared} old value(s) from an earlier call were removed from the empty lines of this table.")
    for c in t.unmatched:
        res.unrecognised.append(t.texts[c][-1].strip().lstrip("*# ").strip())
    res.missing.extend(sorted(yellow_missing))
    return len(records)


# ------------------------------------------------------------------ single fields (label / value)
def find_value_cell(ws, r, c, force_below=False):
    """Where does the value for the label at (r, c) go? Returns (cell, looks_like_input, already_filled)."""
    c0, c1, r0, r1 = span(ws, r, c)
    cands = []
    for col in (range(c1 + 1, min(c1 + 7, ws.max_column + 1)) if not force_below else []):
        a = anchor(ws, r0, col)
        if not blank(a.value):
            if is_formula(a.value) or not isinstance(a.value, str):
                return a, True, True
            break
        looks = (not a.protection.locked) or has_dv(ws, a) or has_border(a)
        if looks:
            return a, True, False
        cands.append(a)
    below = anchor(ws, r1 + 1, c0)
    if blank(below.value):
        return below, has_border(below) or (not below.protection.locked), False
    if not force_below and not blank(below.value) and (is_formula(below.value) or not isinstance(below.value, str)):
        return below, True, True
    if cands:
        return cands[0], False, False
    return None, False, False


def fill_fields(ctx, ws, res, allowed_fields, exclude, section_aware=False):
    master, wb = ctx["master"], ctx["wb"]
    found = []
    section = None
    for r in range(1, min(ws.max_row, 60) + 1):
        for c in range(1, min(ws.max_column, 12) + 1):
            if (r, c) in exclude:
                continue
            v = ws.cell(r, c).value
            if not isinstance(v, str) or not v.strip() or len(v) > 60 or isinstance(ws.cell(r, c), MergedCell):
                continue
            n = norm(v)
            if section_aware:
                if n.startswith("ship company"):
                    section = "company"
                elif n.startswith("company security officer"):
                    section = "cso"
                elif n.startswith("port state control"):
                    section = "psc"
            concept, s = match(v, V.FIELDS, allowed=allowed_fields, exclude=V.FIELD_EXCLUDE)
            if concept:
                if concept == "imo" and section == "company":
                    concept = "company_imo"
                found.append((r, c, concept))
    results, cols = [], []
    for r, c, concept in found:
        cell, looks, filled = find_value_cell(ws, r, c)
        results.append([r, c, concept, cell, looks, filled])
    # labels that sit side by side in one row with their values underneath (e.g. NSW forms): use 'below' for all
    below_rows = {r for r, c, concept, cell, looks, filled in results
                  if cell is not None and cell.row > span(ws, r, c)[3] and not filled}
    for item in results:
        r, c = item[0], item[1]
        if r in below_rows and item[3] is not None and item[3].row == span(ws, r, c)[2]:
            cell, looks, filled = find_value_cell(ws, r, c, force_below=True)
            if cell is not None and looks:
                item[3], item[4], item[5] = cell, looks, filled
    for r, c, concept, cell, looks, filled in results:
        if cell is not None and looks and not filled:
            cols.append(cell.column)
    consensus = max(set(cols), key=cols.count) if cols else None
    count = 0
    for r, c, concept, cell, looks, filled in results:
        if filled:
            continue
        if cell is None or not looks:
            if consensus:
                a = anchor(ws, r, consensus)
                if blank(a.value):
                    cell = a
        if cell is None:
            cell = ws.cell(r, c)
            mark(cell)
            res.yellow += 1
            continue
        label = ws.cell(r, c).value.strip()
        if concept in V.UNSURE:
            mark(cell)
            res.yellow += 1
            res.note(f"'{label}': not filled, unclear how this form wants it - please complete by hand.")
            continue
        if concept in V.NOT_IN_MASTER:
            mark(cell)
            res.yellow += 1
            res.missing.append(label)
            continue
        key = "w_next_port" if concept == "w_next_port_hdr" else concept
        raw = master.fields.get(key) if key in master.fields else master.waste["header"].get(key)
        if blank(raw):
            mark(cell)
            res.yellow += 1
            res.missing.append(label)
            continue
        if concept in ("issc_expiry", "w_last_date", "eta"):
            d = C.to_date(raw)
            if d:
                if not C.is_date_format(cell.number_format):
                    cell.number_format = "dd-mm-yyyy"
                cell.value = d
                count += 1
                continue
        allowed = get_allowed(wb, ws, cell)
        if allowed:
            pick = C.choose_from_list(raw, allowed)
            if pick is None:
                cell.value = raw
                mark(cell)
                res.yellow += 1
                res.note(f"Value for '{label}' is not in the form's list (master value copied, please check).")
                continue
            raw = pick
        cell.value = raw
        count += 1
    return count


# ------------------------------------------------------------------ waste
def waste_code(text):
    m = re.search(r"\b(\d{3})\b", str(text))
    return m.group(1) if m else None

def fill_waste(ctx, ws, t, res):
    master = ctx["master"]
    rows = master.waste["rows"]
    wcols = {c: k for c, k in t.cols.items()}
    label_col = next((c for c, k in wcols.items() if k == "w_type"), 1)
    value_cols = {c: k for c, k in wcols.items() if k != "w_type"}
    # fixed layout (labels pre-printed in the form) or free layout (type chosen from a list)?
    labels = []
    for r in range(t.data_start, min(ws.max_row, t.data_start + 80) + 1):
        v = ws.cell(r, label_col).value
        if isinstance(v, str) and v.strip():
            c0, c1, _, _ = span(ws, r, label_col)
            if c1 >= max(value_cols):   # a merged group title such as '1. Engine room'
                continue
            labels.append((r, v))
    count = 0
    mast_next = master.waste["header"].get("w_next_port")
    filled_rows = [x for x in rows if x["values"]]
    if len(labels) >= 3:   # fixed layout
        used = set()
        for r, label in labels:
            code = waste_code(label)
            target = None
            for i, m in enumerate(filled_rows):
                if i in used:
                    continue
                if code and m["code"] == code or (not code and norm(m["name"]) == norm(label) and not m["code"]):
                    target = (i, m)
                    break
            if not target:
                continue
            used.add(target[0])
            vals = dict(target[1]["values"])
            vals["w_next_port"] = mast_next
            for c, k in value_cols.items():
                a = anchor(ws, r, c)
                raw = vals.get(k)
                if blank(raw):
                    continue
                a.value = raw
            count += 1
        for i, m in enumerate(filled_rows):
            if i not in used:
                res.note(f"Master waste line '{m['name']}' has no matching line on this form; not filled.")
        return count
    # free layout: one line per waste item, type chosen from a list
    types = []
    for other in ctx["wb"].worksheets:
        if other is ws:
            continue
        for c in range(1, 4):
            if norm(other.cell(1, c).value or "") in ("waste types", "waste type"):
                types = [other.cell(r, c).value for r in range(2, other.max_row + 1) if other.cell(r, c).value]
    records = []
    for m in filled_rows:
        pick = None
        for tp in types:
            if m["code"] and waste_code(tp) == m["code"]:
                pick = tp.strip()
                break
        rec = dict(m["values"])
        rec["w_type"] = pick or f"{m['name']} {m['code']}".strip()
        rec["w_next_port"] = mast_next
        records.append(rec)
    if not records:
        return 0
    tt = t
    tt.capacity = max(0, tt.capacity)
    if len(records) > tt.capacity:
        extend_table(ws, tt, len(records) - tt.capacity, res)
    for i, rec in enumerate(records):
        row = tt.data_start + i
        for c, k in wcols.items():
            a = anchor(ws, row, c)
            if not blank(rec.get(k)):
                a.value = rec[k]
        count += 1
    return count


# ------------------------------------------------------------------ ship's stores
def store_class(text):
    n = " " + norm(text) + " "
    best, bl = None, 0
    for cls, kws in V.STORE_CLASSES:
        for kw in kws:
            if f" {kw} " in n or (len(kw) > 3 and kw in n):
                if len(kw) > bl:
                    best, bl = cls, len(kw)
    return best

UNIT_FAMILY = {
    "l": ("volume", 1), "ltr": ("volume", 1), "liter": ("volume", 1), "liters": ("volume", 1), "litre": ("volume", 1),
    "litres": ("volume", 1), "lt": ("volume", 1), "m3": ("volume", 1000), "cl": ("volume", 0.01), "ml": ("volume", 0.001),
    "kg": ("mass", 1), "kgs": ("mass", 1), "kilo": ("mass", 1), "kilos": ("mass", 1), "kilogram": ("mass", 1),
    "t": ("mass", 1000), "mt": ("mass", 1000), "tonnes": ("mass", 1000), "tons": ("mass", 1000),
    "pcs": ("count", 1), "pc": ("count", 1), "pieces": ("count", 1), "piece": ("count", 1), "sticks": ("count", 1),
}

def fill_stores(ctx, ws, res):
    master = ctx["master"]
    arts = []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and re.match(r"^\s*\d{3}\s*-\s*\S", v):
            arts.append((r, v))
    if len(arts) < 4:
        return None
    qty_col = unit_col = None
    for r in range(1, arts[0][0]):
        for c in range(1, ws.max_column + 1):
            n = norm(ws.cell(r, c).value or "")
            if n in ("quantity", "menge", "cantidad"):
                qty_col = c
            if n in ("unit", "einheit", "unidad"):
                unit_col = c
    if not qty_col:
        return None
    classes = {}
    for it in master.stores:
        cls = store_class(it["name"]) or store_class(it["category"])
        if not cls and "alcohol" in norm(it["category"]):
            cls = "other_alcohol"
        classes.setdefault(cls, []).append(it)
    count = 0
    for r, label in arts:
        cls = store_class(re.sub(r"^\s*\d{3}\s*-\s*", "", label))
        items = classes.get(cls) if cls else None
        if not items:
            continue
        unit_text = ws.cell(r, unit_col).value if unit_col else None
        fam = UNIT_FAMILY.get(compact(unit_text or ""))
        total, ok = 0.0, True
        for it in items:
            f = UNIT_FAMILY.get(compact(it["unit"]))
            try:
                q = float(it["qty"])
            except (TypeError, ValueError):
                ok = False
                continue
            if fam and f and f[0] == fam[0]:
                total += q * f[1] / fam[1]
            elif not fam and not f and norm(it["unit"]) == norm(unit_text or ""):
                total += q
            else:
                ok = False
        a = anchor(ws, r, qty_col)
        if ok:
            a.value = int(total) if float(total).is_integer() else round(total, 2)
            count += 1
        else:
            mark(a)
            res.yellow += 1
            res.note(f"'{label.strip()}': the unit in the master ({', '.join(sorted({str(i['unit']) for i in items}))}) "
                     f"cannot be converted safely to {unit_text}; left empty, please fill in by hand.")
    return count


# ------------------------------------------------------------------ one sheet
PERSON_MIN = 4

def process_sheet(ctx, ws):
    res = SheetResult(ws.title)
    if re.search(r"instruction|reference|unlocode|locode|notice|conditions|taul|^list$", ws.title, re.I):
        res.note("Information sheet; left untouched.")
        return res
    master = ctx["master"]
    exclude = set()
    is_mdoh = H.is_health(ws)
    title_text = norm(ws.title + " " + " ".join(str(ws.cell(r, c).value) for r in range(1, 9)
                                                for c in range(1, 6) if ws.cell(r, c).value))
    # 1) ship's stores
    n_art = sum(1 for r in range(1, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str)
                and re.match(r"^\s*\d{3}\s*-\s*\S", ws.cell(r, 1).value))
    has_qty = any(norm(ws.cell(r, c).value or "") in ("quantity", "menge", "cantidad")
                  for r in range(1, 12) for c in range(1, 12))
    if n_art >= 4 and has_qty:
        res.kind = "Ship's stores"
        n = fill_stores(ctx, ws, res)
        fields = fill_fields(ctx, ws, res, V.HEADER_FIELDS, exclude)
        res.lines.append(f"{n or 0} article {'line' if n == 1 else 'lines'} filled")
        return res
    # 1b) article list with printed article names (FAL 3 / provisions)
    art_cols = ST.find_article_table(ws)
    if art_cols and not is_mdoh:
        ST.fill_stores_sheet(ctx, ws, res, art_cols)
        return res
    # 2) tables
    cand = []
    for kind, vocab in (("waste", V.WASTE), ("ports", V.PORTS), ("person", V.PERSON)):
        for t in detect_tables(ws, vocab):
            n_concepts = len(set(t.cols.values()) - {"serial"})
            cand.append((n_concepts, kind, t))
    # keep only the best interpretation per header row
    chosen, seen = [], set()
    for n_concepts, kind, t in sorted(cand, key=lambda x: -x[0]):
        if any(r in seen for r in t.rows):
            continue
        chosen.append((kind, t))
        seen.update(t.rows)
    tables = [(k, t) for k, t in chosen if (k != "person" or len(set(t.cols.values())) >= PERSON_MIN)]
    is_waste2 = (not is_mdoh) and not any(k == "waste" for k, _ in tables) and W2.is_waste_sheet(ws)
    if is_mdoh:
        tables = [(k, t) for k, t in tables if k == "ports"]
        ctx["has_ports_table"] = bool(tables)
    if is_waste2 and W2.fill_waste_sheet(ctx, ws, res):
        return res
    for kind, t in tables:
        for r in t.rows:
            for c in range(1, ws.max_column + 1):
                exclude.add((r, c))
        for r in range(t.data_start, t.data_start + max(t.capacity, 1) + 80):
            for c in range(1, ws.max_column + 1):
                exclude.add((r, c))
    kinds = []
    for kind, t in sorted(tables, key=lambda x: x[1].rows[0]):
        if kind == "person":
            is_pax = ("passenger" in title_text or "passagier" in title_text or "pasajero" in title_text) \
                     and "crew" not in title_text
            recs = master.pax if is_pax else master.crew
            res.kind = "Passenger list" if is_pax else "Crew list"
            n = fill_table(ctx, ws, t, recs, res)
            res.lines.append(f"{n} of {len(recs)} {'line' if len(recs) == 1 else 'lines'} filled" if recs else "no data in the master, nothing to fill")
        elif kind == "ports":
            nextrow = "next_port" in t.cols.values()
            if nextrow:
                recs = [{"next_port": master.fields.get("next_port")}] if master.fields.get("next_port") else []
                res.kind = "Port calls"
                n = fill_table(ctx, ws, t, recs, res)
                res.lines.append("next port filled" if n else "no next port in the master")
            else:
                recs = master.ports
                res.kind = "Last 10 ports"
                n = fill_table(ctx, ws, t, recs[:10], res)
                res.lines.append(f"{n} of {min(len(recs), 10)} port calls filled" if recs else "no port calls in the master")
        elif kind == "waste":
            res.kind = "Waste"
            n = fill_waste(ctx, ws, t, res)
            res.lines.append(f"{n} waste {'line' if n == 1 else 'lines'} filled")
    if is_mdoh:
        H.fill_health(ctx, ws, res, exclude)
        return res
    # 3) single fields
    sec_hits = 0
    if not tables:
        for r in range(1, min(ws.max_row, 60) + 1):
            for c in range(1, min(ws.max_column, 12) + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str) and match(v, V.FIELDS, allowed=V.SECURITY_FIELDS)[0]:
                    sec_hits += 1
    if sec_hits >= 4:
        res.kind = "Security information"
        n = fill_fields(ctx, ws, res, V.SECURITY_FIELDS | V.HEADER_FIELDS, exclude, section_aware=True)
        res.lines.append(f"{n} {'field' if n == 1 else 'fields'} filled")
    else:
        allowed = set(V.HEADER_FIELDS)
        if any(k == "waste" for k, _ in tables):
            allowed |= V.WASTE_HEADER_FIELDS
        n = fill_fields(ctx, ws, res, allowed, exclude)
        if res.kind != "not recognised" and n:
            res.lines.append(f"{n} header {'field' if n == 1 else 'fields'} filled")
    return res


def process_workbook(master, path, out_dir):
    import os
    wb = load_workbook(path)
    ctx = {"master": master, "wb": wb, "rules": parse_instructions(wb), "path": path}
    results = []
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        results.append(process_sheet(ctx, ws))
    out_path = os.path.join(out_dir, os.path.basename(path))
    os.makedirs(out_dir, exist_ok=True)
    wb.save(out_path)
    return out_path, results, package_notes(path)

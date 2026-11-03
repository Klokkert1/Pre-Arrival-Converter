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
import msw as MSW
import fal1 as F1
import fal4 as F4
import poc as POC
import nmsw as NM
import shipdetails as SD
import cers3 as CE
import dakosy as DK
import noad as NA

# openpyxl recomputes max_row / max_column over ALL cells at every call; on big templates (5000 empty
# formatted rows) the program called them thousands of times. They are cached per sheet and reset when rows are added.
from openpyxl.worksheet.worksheet import Worksheet as _WS

def _cached_dim(name):
    orig = getattr(_WS, name).fget
    def getter(self):
        hit = self.__dict__.get("_pf_" + name)
        if hit is not None:
            return hit
        val = orig(self)
        self.__dict__["_pf_" + name] = val
        return val
    return property(getter)

if not getattr(_WS, "_pf_patched", False):
    _WS.max_row = _cached_dim("max_row")
    _WS.max_column = _cached_dim("max_column")
    _WS._pf_patched = True

def reset_dims(ws):
    ws.__dict__.pop("_pf_max_row", None)
    ws.__dict__.pop("_pf_max_column", None)

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
    for r in range(1, min(ws.max_row, 75) + 1):
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
            elif concept == "port":
                t.cols[c] = "port_name"          # a second port column (e.g. 'Port LOCODE' and 'Port name')
        # one header 'date of arrival / departure' above two or three narrow columns: first = arrival, last = departure
        for c in list(t.cols):
            if t.cols[c] in ("arrival", "departure") and "arrival" in vocab and "departure" in vocab:
                tx = norm(" ".join(texts.get(c, [])))
                if re.search(r"arriv", tx) and re.search(r"depart", tx):
                    extra = []
                    cc = c + 1
                    while cc <= c + 2 and cc not in texts and cc not in t.cols and cc <= ws.max_column:
                        extra.append(cc)
                        cc += 1
                    if extra:
                        t.cols[c] = "arrival"
                        t.cols[extra[-1]] = "departure"
                        for _mid in extra[:-1]:
                            t.cols[_mid] = "transit_time"
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
        # guidance line(s) under the header ("Enter the 3-letter country code...", "Only values from list allowed")
        for _ in range(2):
            _g = [ws.cell(t.data_start, c).value for c in keycols
                  if not isinstance(ws.cell(t.data_start, c), MergedCell)]
            _g = [x for x in _g if isinstance(x, str) and x.strip()]
            if len(_g) >= 2 and sum(1 for x in _g if re.match(r"(?i)\s*(enter|only|please|select|use|insert|fill|choose|example|e\.g\.|max|must|if )", x)
                                    or len(x) > 60) >= max(2, len(_g) // 2):
                t.data_start += 1
        if sc:
            # a blank spacer line between the header and the numbered lines
            for _ in range(2):
                _row = [ws.cell(t.data_start, c).value for c in range(1, min(ws.max_column, 20) + 1)
                        if not isinstance(ws.cell(t.data_start, c), MergedCell)]
                _nx = ws.cell(t.data_start + 1, sc).value
                if all(blank(x) for x in _row) and not blank(_nx) and re.fullmatch(r"\d{1,3}", str(_nx).strip()):
                    t.data_start += 1
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
    reset_dims(ws)
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
    if concept in ("port", "port_name"):
        want_locode = "locode" in norm(hints) or MSW.is_msw_book(ctx["wb"])
        raw = (rec.get("locode") or rec.get("port")) if want_locode else rec.get("port")
        return (raw, "ok") if not blank(raw) else (None, "missing" if mandatory else "blank")
    raw = rec.get(concept)
    if concept == "row_ship_name":
        raw = re.sub(r"^\s*m\.?\s?v\.?\s+", "", str(ctx["master"].fields.get("ship_name") or ""), flags=re.I) or None
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
    _old = cell.value if isinstance(cell.value, str) else None
    _exs = ctx.setdefault("col_examples", {})
    _key = (ws.title, cell.column, concept)
    if _old and re.fullmatch(r"[A-Z]{2,3}|Male|Female|M|F|MALE|FEMALE|male|female", _old.strip()):
        _exs.setdefault(_key, _old)
    elif not _old and _key in _exs:
        _old = _exs[_key]
    if _old and not allowed:                      # the old value shows which form the box wants
        _o = _old.strip()
        if concept in COUNTRY_CONCEPTS:
            if re.fullmatch(r"[A-Z]{2}", _o):
                hints = hints + " iso 2 code"
            elif re.fullmatch(r"[A-Z]{3}", _o):
                hints = hints + " iso 3"
        elif concept == "gender":
            if _o.lower() in ("male", "female"):
                hints = hints + " male female"
            elif _o.upper() in ("M", "F"):
                hints = hints + " m f"
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


ID_SET = {"id_type", "id_number", "id_issue_country", "id_expiry", "id_issue_date", "issuer_entity"}


def doc_rows(ws, t, records):
    """Forms (e.g. IMO FAL 5) that show a second line per person with only the identity-document fields:
    when the form's own example data has such lines, each person gets one line per document."""
    pcols = [c for c, k in t.cols.items() if k in ("family_name", "given_names")]
    dcols = [c for c, k in t.cols.items() if k in ("id_number", "id_type")]
    if not pcols or not dcols:
        return records, False
    cont = False
    for row in range(t.data_start + 1, t.data_start + min(t.capacity, 60)):
        if all(blank(ws.cell(row, c).value) for c in pcols) and any(not blank(ws.cell(row, c).value) for c in dcols) \
                and not any(is_formula(ws.cell(row, c).value) for c in dcols) \
                and not blank(ws.cell(row - 1, pcols[0]).value):
            cont = True
            break
    scol = [c for c, k in t.cols.items() if k == "serial"]
    if not cont and scol and dcols:      # blank form: numbers printed on every second line, document codes (P / SB) printed
        r0 = t.data_start
        if not blank(ws.cell(r0, scol[0]).value) and blank(ws.cell(r0 + 1, scol[0]).value) \
                and not blank(ws.cell(r0 + 2, scol[0]).value) and blank(ws.cell(r0 + 3, scol[0]).value):
            cont = True
    if not cont:
        return records, False
    out = []
    for r in records:
        first = dict(r)
        out.append(first)
        if r.get("pp_no") and r.get("sb_no"):
            out.append({"_cont": True, "id_type": "Seaman's book", "id_number": r["sb_no"],
                        "id_issue_country": r.get("sb_state"), "id_expiry": r.get("sb_exp")})
            first.update({"id_type": "Passport", "id_number": r["pp_no"], "id_issue_country": r.get("pp_state"),
                          "id_expiry": r.get("pp_exp")})
            for k in [k for k, v in first.items() if v is None]:
                del first[k]
        else:
            out.append({"_cont": True})      # keep the two-line layout; second line stays empty
    return out, True


def fill_table(ctx, ws, t, records, res):
    n_people = len(records)
    ctx["pair_mode"] = False
    if records and t.cols and "id_number" in t.cols.values() and ctx.get("crew_recs") is records:
        records, paired = doc_rows(ws, t, records)
        ctx["pair_mode"] = paired
        if paired:
            res.note("Passport and seaman's book are entered on two lines per person, as this form shows.")
    # skip tables that are driven by formulas
    probe = [ws.cell(t.data_start, c).value for c in t.cols]
    if probe and sum(1 for v in probe if is_formula(v)) * 2 >= len(probe):
        res.note("This sheet is filled by formulas from another sheet; left untouched.")
        return 0
    if len(records) > t.capacity:
        extend_table(ws, t, len(records) - t.capacity, res)
    yellow_missing = set()
    person_no = 0
    for i, rec in enumerate(records):
        row = t.data_start + i
        cont = bool(rec.get("_cont"))
        if not cont:
            person_no += 1
        for c, concept in t.cols.items():
            cell = ws.cell(row, c)
            if cont and concept not in ID_SET:
                if not isinstance(cell, MergedCell) and not blank(cell.value) and not is_formula(cell.value):
                    cell.value = None
                continue
            if isinstance(cell, MergedCell):
                cell = anchor(ws, row, c)
            if concept == "serial":
                if blank(cell.value):
                    cell.value = person_no
                continue
            rule = rule_for(ctx, t.texts.get(c, []))
            mand = c in t.mandatory or bool(re.search(r"is mandatory\.", rule, re.I))
            if concept == "id_type" and ctx.get("pair_mode") and not blank(cell.value) and not is_formula(cell.value):
                continue          # document code (P / SB) already printed on the form
            val, status = convert_cell(ctx, ws, cell, concept, rec, t.hints.get(c, ""), mand, rule)
            if cont and status in ("blank", "missing"):
                if not blank(cell.value) and not is_formula(cell.value):
                    cell.value = None
                continue
            if status == "ok" or (status == "unsure" and val is not None):
                if blank(cell.value) or not is_formula(cell.value):
                    cell.value = val
            elif status in ("blank", "missing") and not blank(cell.value) and not is_formula(cell.value) \
                    and concept not in ("serial",):
                cell.value = None                       # old value of an earlier voyage; the master has nothing here
                status = "missing"
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
    _end = t.data_start + t.capacity
    while _end < min(ws.max_row, t.data_start + 400) and any(
            not isinstance(ws.cell(_end, c), MergedCell) and not blank(ws.cell(_end, c).value)
            and not is_formula(ws.cell(_end, c).value) for c in t.cols if t.cols[c] != "serial"):
        _end += 1                                   # leftover filled lines below the printed capacity
    for row in range(t.data_start + len(records), _end):
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
    return n_people


# ------------------------------------------------------------------ single fields (label / value)
def _stale_value(ws, a):
    """A text cell right of a label: an old filled-in value (True) or the next label (False)?"""
    t = str(a.value).strip()
    if not t or t.endswith(":") or len(t) > 90 or (len(t.split()) > 3 and not re.search(r"\d", t)):
        return False
    if a.font is not None and a.font.b:
        return False
    if not (has_border(a) or not a.protection.locked or has_dv(ws, a)):
        return False
    for vocab in (V.FIELDS, V.PERSON, V.PORTS, V.WASTE):
        if match(t, vocab)[0]:
            return False
    return True


def find_value_cell(ws, r, c, force_below=False):
    """Where does the value for the label at (r, c) go? Returns (cell, looks_like_input, already_filled)."""
    c0, c1, r0, r1 = span(ws, r, c)
    cands = []
    _start = c1 + 1
    for _col in range(c1 + 1, min(c1 + 4, ws.max_column + 1)):
        _sv = anchor(ws, r0, _col).value
        if isinstance(_sv, str) and _sv.strip() in (":", "=", ":-"):
            _start = _col + 1            # the value comes after the separator
            break
    for col in (range(_start, min(_start + 6, ws.max_column + 1)) if not force_below else []):
        a = anchor(ws, r0, col)
        if isinstance(a.value, str) and a.value.strip() in (":", "=", "-", "–", ":-", "->"):
            continue                       # a separator between label and value
        if not blank(a.value):
            if is_formula(a.value):
                return a, True, True
            if not isinstance(a.value, str):
                if col == _start and _start > c1 + 1:
                    return a, True, False
                if has_border(a) or not a.protection.locked or has_dv(ws, a):
                    return a, True, False      # an old number/date from an earlier call: will be replaced
                return a, True, True
            if col == _start and _start > c1 + 1 and not a.value.strip().endswith(":"):
                return a, True, False          # first cell after a ':' separator is the value (old or empty)
            if _stale_value(ws, a):
                return a, True, False          # an old value from an earlier call: will be replaced
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
        for c in range(1, min(ws.max_column, 20) + 1):
            if (r, c) in exclude:
                continue
            v = ws.cell(r, c).value
            if not isinstance(v, str) or not v.strip() or len(v) > 90 or isinstance(ws.cell(r, c), MergedCell):
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
        if concept == "eta_dt":
            raw = master.fields.get("eta_arrival") or master.fields.get("eta")
        if concept == "etd_dt":
            raw = master.fields.get("etd_departure")
        if concept == "draft_fa":
            _fo, _af = master.fields.get("draft_fore"), master.fields.get("draft_aft")
            raw = f"F= {_fo}  A= {_af}" if not blank(_fo) and not blank(_af) else None
        if concept in ("port_mv", "port_arr", "port_dep", "date_mv", "date_arr", "date_dep"):
            _f = master.fields
            _dep = str(_f.get("movement") or "Arrival").strip().lower().startswith("dep")
            if concept == "port_mv" or concept == "port_arr" or (concept == "port_dep" and _dep):
                raw = _f.get("port")
            elif concept == "port_dep":
                raw = _f.get("last_port")
            elif concept == "date_arr" or (concept == "date_mv" and not _dep):
                raw = _f.get("eta_arrival") or _f.get("eta")
            else:
                raw = _f.get("etd_departure")
        if blank(raw):
            mark(cell)
            res.yellow += 1
            res.missing.append(label)
            continue
        if concept in ("eta_dt", "etd_dt"):
            if hasattr(raw, "hour"):
                if not C.is_date_format(cell.number_format):
                    cell.number_format = "dd-mm-yyyy hh:mm"
                cell.value = raw
                count += 1
                continue
        if concept in ("issc_expiry", "w_last_date", "eta", "date_mv", "date_arr", "date_dep"):
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
    # other languages
    "liter": ("volume", 1), "litros": ("volume", 1), "litro": ("volume", 1), "litri": ("volume", 1), "litro": ("volume", 1),
    "litr": ("volume", 1), "litry": ("volume", 1), "litrov": ("volume", 1), "liters": ("volume", 1), "l": ("volume", 1),
    "kilo": ("mass", 1), "kilos": ("mass", 1), "kilogramm": ("mass", 1), "kilogramos": ("mass", 1), "kilogrammes": ("mass", 1),
    "tonne": ("mass", 1000), "tonnen": ("mass", 1000), "toneladas": ("mass", 1000), "tonnellate": ("mass", 1000),
    "stuks": ("count", 1), "stuk": ("count", 1), "stk": ("count", 1), "stck": ("count", 1), "stuck": ("count", 1),
    "pieces": ("count", 1), "pieza": ("count", 1), "piezas": ("count", 1), "pezzi": ("count", 1), "pecas": ("count", 1),
    "unidades": ("count", 1), "unites": ("count", 1), "nr": ("count", 1), "szt": ("count", 1), "kpl": ("count", 1),
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
    reset_dims(ws)
    res = SheetResult(ws.title)
    if re.fullmatch(r"\s*list\s*", ws.title, re.I) and ws.max_column <= 3:
        res.note("Information sheet; left untouched.")
        return res
    if re.search(r"instruction|reference|unlocode|locode|notice|conditions|taul|document type|dictionary|country codes", ws.title, re.I):
        res.note("Information sheet; left untouched.")
        return res
    master = ctx["master"]
    if re.search(r"health.*page\s*[2-9]|declaration of health.*page\s*[2-9]", ws.title, re.I):
        res.note("Continuation page of the health declaration (persons who joined the ship, port details): "
                 "not filled - complete by hand if it applies.")
        return res
    exclude = set()
    is_mdoh = H.is_health(ws)
    title_text = norm(ws.title + " " + " ".join(str(ws.cell(r, c).value) for r in range(1, 9)
                                                for c in range(1, 6) if ws.cell(r, c).value))
    # 0) MSW upload templates
    if MSW.is_msw_book(ctx["wb"]):
        k = MSW.key(ws.title)
        if k == "wastereceipt":
            res.note("Waste receipt is completed after the port call; left untouched.")
            return res
        if k == "stowawaylist":
            res.kind = "Stowaway list"
            res.note("Stowaway list: not filled (the crew are not stowaways). Fill in by hand if there are stowaways.")
            return res
        if k == "wasteandresidues" and W2.fill_waste_sheet(ctx, ws, res):
            return res
        if k in MSW.MSW_SHEETS and k != "security":
            if MSW.fill_msw(ctx, ws, res):
                return res
        elif k == "security":
            MSW.fill_msw(ctx, ws, res)
            exclude |= {(r, c) for r in range(1, 30) for c in range(1, 12) if ws.cell(r, c).value is not None and r < 14}
    _name = ctx["master"].fields.get("ship_name")
    if _name:
        _rx = re.compile(r'(m\.?v\.?\s*)"?\s*ship name\s*"?', re.I)
        for (_r, _c), _cell in list(ws._cells.items()):
            if isinstance(_cell.value, str) and not _cell.value.startswith("=") and _rx.search(_cell.value):
                _cell.value = _rx.sub(lambda mm: mm.group(1) + re.sub(r"^\s*m\.?\s?v\.?\s+", "", str(_name), flags=re.I), _cell.value)
    if DK.is_dakosy(ctx["wb"]) and ws.title.lower() not in ("instructions", "f1", "f2"):
        t_ = ws.title.lower()
        if t_ == "waste delivery receipt":
            res.note("Waste delivery receipt is completed after the port call; left untouched.")
            return res
        if t_ in ("main", "ship", "port message", "traffic message on arrival", "traffic message on departure",
                  "security", "health", "waste", "stowaways"):
            DK.fill_dakosy(ctx, ws, res)
            if t_ == "waste":
                tb = {"row": 17, "cols": {5: "w_deliver", 6: "w_capacity", 7: "w_onboard", 8: "w_next_port",
                                          9: "w_generated"}, "label": 2, "alt": 2}
                _h = ctx["master"].waste["header"]
                _old = _h.get("w_next_port")
                if ctx["master"].fields.get("next_port_locode"):
                    _h["w_next_port"] = ctx["master"].fields["next_port_locode"]      # DAKOSY wants the LOCODE
                try:
                    W2.fill_table(ctx, ws, res, tb, keep_unmatched=True)
                finally:
                    _h["w_next_port"] = _old
                res.kind = "DAKOSY waste"
                return res
            if t_ in ("main", "ship", "port message", "traffic message on arrival", "traffic message on departure", "health", "stowaways"):
                return res
            exclude |= {(_r, _c) for _r in range(1, 28) for _c in range(1, 12)}
    if NA.is_noad(ctx["wb"]) and re.match(r"^[1-5]\. ", ws.title):
        NA.fill_noad(ctx, ws, res)
        if not ws.title.startswith("1."):
            return res
        exclude |= {(_r, _c) for _r in range(1, 51) for _c in range(1, 18)}
    _cers = CE.is_cers(ctx["wb"]) and ws.title.lower() in ("vessel", "voyage", "waste", "security")
    if _cers:
        CE.fill_cers(ctx, ws, res)
        if ws.title.lower() in ("vessel", "voyage", "waste"):
            return res
        exclude |= {(_r, _c) for _r in range(1, 34) for _c in range(1, 16)}
    if SD.is_shipdetails(ws) and SD.fill_shipdetails(ctx, ws, res):
        return res
    if NM.is_nmsw(ws) and NM.fill_nmsw(ctx, ws, res):
        return res
    if F1.is_fal1(ws) and F1.fill_fal1(ctx, ws, res):
        return res
    if POC.is_poc(ws) and POC.fill_poc(ctx, ws, res):
        return res
    if F4.is_templist(ws) and F4.fill_templist(ctx, ws, res):
        return res
    if F4.fill_fal4(ctx, ws, res):
        return res
    # 1) ship's stores
    n_art = sum(1 for (r, c), cell in list(ws._cells.items())
                if c == 1 and isinstance(cell.value, str) and re.match(r"^\s*\d{3}\s*-\s*\S", cell.value))
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
            ctx["crew_recs"] = recs if not is_pax else None
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
        if res.kind == "not recognised" and n >= 2 and not tables:
            res.kind = "Information form"
        if res.kind != "not recognised" and n:
            res.lines.append(f"{n} header {'field' if n == 1 else 'fields'} filled")
    return res


def _static_language_labels(wb, path):
    """Labels that are language-switch formulas (=IF(Settings!B6="bg","..","Store item name")) become their
    displayed text, so the label recognition can read them."""
    if not any("Settings!" in str(c.value) for ws in wb.worksheets[:3] for row in ws.iter_rows(min_row=1, max_row=4)
               for c in row if isinstance(c.value, str) and c.value.startswith("=")):
        return
    wv = load_workbook(path, data_only=True)
    for ws in wb.worksheets:
        if ws.title not in wv.sheetnames:
            continue
        wsv = wv[ws.title]
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 60)):
            for c in row:
                f = c.value
                if (isinstance(f, str) and f.startswith("=") and "Settings!" in f and "_N_" not in f
                        and not isinstance(c, MergedCell)):
                    v = wsv[c.coordinate].value
                    if isinstance(v, str) and v.strip() and not v.startswith("DO NOT FILL"):
                        c.value = v.strip()


def process_workbook(master, path, out_dir):
    import os
    import slim as _slim
    _wpath, _tmp = _slim.slim(path)
    try:
        wb = load_workbook(_wpath)
    finally:
        if _tmp:
            try:
                os.remove(_wpath)
            except OSError:
                pass
    try:
        _static_language_labels(wb, _wpath if os.path.exists(_wpath) else path)
    except Exception:
        pass
    ctx = {"master": master, "wb": wb, "rules": parse_instructions(wb), "path": path}
    results = []
    try:
        import extras as _EXS
        _labels = _EXS.snapshot_labels(wb)
    except Exception:
        _EXS, _labels = None, {}
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        _r = process_sheet(ctx, ws)
        try:
            import extras as EX
            EX.post(ctx, ws, _r)
        except Exception as _e:
            _r.note(f'extra header step skipped: {_e!r}'[:120])
        try:
            if _EXS and _r.kind and "not recognised" not in str(_r.kind):
                _EXS.report_unfilled(ws, _r, _labels.get(ws.title, {}))
        except Exception:
            pass
        results.append(_r)
    out_path = os.path.join(out_dir, os.path.basename(path))
    os.makedirs(out_dir, exist_ok=True)
    wb.save(out_path)
    notes = package_notes(path)
    try:
        _mn = re.sub(r"\W+", " ", str(master.fields.get("master_name") or "")).lower().split()
        _mn = [w for w in _mn if w not in ("capt", "captain", "master", "kapitan")]
        for _p in master.crew:
            if re.search(r"master|captain|kapit", str(_p.get("rank") or ""), re.I):
                _cn = re.sub(r"\W+", " ", f"{_p.get('given_names') or ''} {_p.get('family_name') or ''}").lower().split()
                if _mn and _cn and not (set(_cn) <= set(_mn) or set(_mn) <= set(_cn)):
                    notes.append(f"Check your master: 'Name of master' in sheet Ship is '{master.fields.get('master_name')}', "
                                 f"but the Crew_List has '{_p.get('given_names')} {_p.get('family_name')}' as Master.")
                break
    except Exception:
        pass
    return out_path, results, notes

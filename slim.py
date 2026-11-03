"""Very large sheets (e.g. 100,000 empty but formatted rows) can use many GB of memory when loaded.
slim() makes a trimmed temporary copy without the empty rows at the bottom of such sheets."""
import os, re, tempfile, zipfile

LIMIT = 1_500_000          # bytes of sheet xml above which we look at the sheet
ROW = re.compile(r"<row\b[^>]*?(?:/>|>.*?</row>)", re.S)
HAS_CONTENT = re.compile(r"<v>|<v [^>]*>|<is>|<f>|<f [^>]*>")


MERGE = re.compile(r'<mergeCell\s+ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\s*/>')


def _col(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


def _huge_merges(xml):
    """Merged ranges covering more than 20,000 cells (e.g. 'grey out the rest of the sheet'): openpyxl would create
    millions of cell objects for them."""
    bad = []
    for m in MERGE.finditer(xml):
        cells = (_col(m.group(3)) - _col(m.group(1)) + 1) * (int(m.group(4)) - int(m.group(2)) + 1)
        if cells > 20000:
            bad.append(m)
    return bad


def _unmerge(xml, bad):
    for m in reversed(bad):
        xml = xml[:m.start()] + xml[m.end():]
    n = len(MERGE.findall(xml))
    if n:
        xml = re.sub(r'<mergeCells count="\d+"', f'<mergeCells count="{n}"', xml, 1)
    else:
        xml = re.sub(r"<mergeCells[^>]*>\s*</mergeCells>", "", xml, 1)
    return xml


def slim(path):
    """Returns (path_to_use, is_temporary)."""
    if not path.lower().endswith((".xlsx", ".xlsm")):
        return path, False
    try:
        with zipfile.ZipFile(path) as z:
            names = [i.filename for i in z.infolist()
                     if i.filename.startswith("xl/worksheets/sheet") and i.filename.endswith(".xml")]
            big = [i.filename for i in z.infolist() if i.filename in names and i.file_size > LIMIT]
            fix = {}
            for n in names:
                x = z.read(n).decode("utf-8", "ignore")
                if "<mergeCell" in x:
                    bad = _huge_merges(x)
                    if bad:
                        fix[n] = bad
            if not big and not fix:
                return path, False
            fd, tmp = tempfile.mkstemp(suffix=os.path.splitext(path)[1])
            os.close(fd)
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
                for info in z.infolist():
                    data = z.read(info.filename)
                    if info.filename in big or info.filename in fix:
                        x = data.decode("utf-8")
                        if info.filename in big:
                            x = _trim(x)
                        if info.filename in fix:
                            x = _unmerge(x, _huge_merges(x))
                        data = x.encode("utf-8")
                    out.writestr(info, data)
            return tmp, True
    except Exception:
        return path, False


def _trim(xml):
    rows = list(ROW.finditer(xml))
    if len(rows) < 300:
        return xml
    last = 0
    for i in range(len(rows) - 1, -1, -1):
        if HAS_CONTENT.search(rows[i].group(0)):
            last = i
            break
    keep = min(len(rows), last + 40)
    if keep >= len(rows):
        return xml
    cut_start = rows[keep].start()
    cut_end = rows[-1].end()
    new = xml[:cut_start] + xml[cut_end:]
    lastrow = re.search(r'\br="(\d+)"', rows[keep - 1].group(0))
    if lastrow:
        new = re.sub(r'<dimension ref="([A-Z]+\d+):([A-Z]+)\d+"\s*/>',
                     lambda m: f'<dimension ref="{m.group(1)}:{m.group(2)}{lastrow.group(1)}"/>', new, 1)
    return new

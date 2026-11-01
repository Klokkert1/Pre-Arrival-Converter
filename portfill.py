#!/usr/bin/env python3
"""PortFill - test version.

Fills port pre-arrival forms (Excel) from the ship's master Excel. The original forms are never changed:
a filled copy is saved in a new folder per port and voyage.

Usage:
  python portfill.py --master Master.xlsx --forms form1.xlsx form2.xlsx [--out OUTPUT_DIR] [--yes]
"""
import argparse, os, sys, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from master import Master
from forms import process_workbook

WARNING = (
    "IMPORTANT: these forms were filled in automatically. You must always check every completed form yourself\n"
    "before it is submitted. The Master / user remains responsible for the content of the forms."
)

def build_report(folder, items):
    lines = [f"Port forms completed - {folder}", ""]
    for name, out_path, results, notes in items:
        lines.append(f"{name}")
        for n in notes:
            lines.append(f"   ! {n}")
        for r in results:
            if r.kind == "not recognised":
                lines.append(f"   - Sheet \"{r.name}\": not recognised, left untouched")
                continue
            detail = "; ".join(r.lines)
            y = f", {r.yellow} cell(s) highlighted yellow" if r.yellow else ""
            lines.append(f"   - {r.kind} (sheet \"{r.name}\"): {detail}{y}")
            if r.unrecognised:
                lines.append(f"       field(s) not recognised: {', '.join(sorted(set(r.unrecognised)))}")
            if r.missing:
                lines.append(f"       not available in the master: {', '.join(sorted(set(r.missing)))}")
            for n in r.notes:
                lines.append(f"       note: {n}")
        lines.append("")
    lines += [WARNING]
    return "\n".join(lines)


def build_html(folder, items):
    """A readable report: file names and sheet names in bold, yellow counts highlighted, warning box."""
    from html import escape as e
    css = ("body{font:15px/1.5 -apple-system,'Segoe UI',Arial,sans-serif;max-width:860px;margin:30px auto;padding:0 20px;color:#1f1f1d}"
           "h1{font-size:22px;font-weight:600;margin:0 0 4px}.sub{color:#5f5e5a;margin:0 0 22px}"
           "h2{font-size:17px;font-weight:700;margin:28px 0 8px;padding:6px 10px;background:#eceae2;border-radius:6px}"
           ".sheet{margin:10px 0 14px 4px;padding-left:12px;border-left:4px solid #c0dd97}"
           ".sheet.warn{border-left-color:#ef9f27}.sheet.none{border-left-color:#d3d1c7;color:#5f5e5a}"
           ".sheet b.name{font-size:15px}.kind{color:#5f5e5a}"
           ".y{background:#ffff00;padding:0 5px;border-radius:3px;font-weight:600}"
           "ul{margin:4px 0 0 0;padding-left:20px;font-size:13.5px;color:#444441}li{margin:2px 0}"
           ".lbl{font-weight:600}.formnote{background:#faeeda;border-radius:6px;padding:6px 10px;margin:6px 0;font-size:13.5px}"
           ".warnbox{margin:32px 0 10px;padding:14px 16px;background:#faeeda;border:2px solid #ef9f27;border-radius:8px;color:#633806}"
           "@media print{body{margin:0}.warnbox{break-inside:avoid}}")
    now = datetime.datetime.now().strftime("%d-%m-%Y %H:%M")
    out = [f"<!DOCTYPE html><html lang='en'><head><meta charset='utf-8'><title>Fill report - {e(folder)}</title>"
           f"<style>{css}</style></head><body>",
           f"<h1>Port forms completed &mdash; {e(folder)}</h1><p class='sub'>Report created {now}</p>"]
    for name, out_path, results, notes in items:
        out.append(f"<h2>{e(name)}</h2>")
        for n in notes:
            out.append(f"<div class='formnote'>{e(n)}</div>")
        for r in results:
            if r.kind == "not recognised":
                if any(x.startswith("Information sheet") for x in r.notes):
                    continue
                out.append(f"<div class='sheet none'><b class='name'>{e(r.name)}</b> &mdash; not recognised, left untouched</div>")
                continue
            sub = []
            if r.unrecognised:
                sub.append(f"<span class='lbl'>Fields not recognised:</span> {e(', '.join(sorted(set(r.unrecognised))))}")
            if r.missing:
                sub.append(f"<span class='lbl'>Not available in the master:</span> {e(', '.join(sorted(set(r.missing))))}")
            for n in r.notes:
                sub.append(e(n))
            formulas_only = (not r.yellow and r.notes and all(x.startswith("This sheet is filled by formulas") for x in r.notes))
            cls = "sheet warn" if (r.yellow or sub) and not formulas_only else "sheet"
            if formulas_only:
                cls = "sheet none"
            y = ""
            if r.yellow:
                y = f" &mdash; <span class='y'>{r.yellow} {'cell' if r.yellow == 1 else 'cells'}</span> marked yellow"
            out.append(f"<div class='{cls}'><b class='name'>{e(r.name)}</b> <span class='kind'>({e(r.kind)})</span>: "
                       f"{e('; '.join(r.lines))}{y}")
            if sub:
                out.append("<ul>" + "".join(f"<li>{x}</li>" for x in sub) + "</ul>")
            out.append("</div>")
    out.append("<div class='warnbox'><b>Important:</b> these forms were filled in automatically. You must always check "
               "every completed form yourself before it is submitted. The Master / user remains responsible for the "
               "content of the forms.</div></body></html>")
    return "".join(out)

def unique_dir(path):
    """Never overwrite an earlier result: add (2), (3) ... when the folder already exists."""
    if not os.path.exists(path):
        return path
    n = 2
    while os.path.exists(f"{path} ({n})"):
        n += 1
    return f"{path} ({n})"

def fill_all(master_path, form_paths, out_root):
    """Fill all forms. Returns a dict with the output folder, the text report and structured results."""
    master = Master(master_path)
    out_dir = unique_dir(os.path.join(out_root, master.folder_name()))
    items = []
    for p in form_paths:
        out_path, results, notes = process_workbook(master, p, out_dir)
        items.append((os.path.basename(p), out_path, results, notes))
    report = build_report(os.path.basename(out_dir), items)
    os.makedirs(out_dir, exist_ok=True)
    with open(os.path.join(out_dir, "Fill report.html"), "w", encoding="utf-8") as f:
        f.write(build_html(os.path.basename(out_dir), items))
    structured = [{
        "name": name, "notes": notes,
        "sheets": [{"name": r.name, "kind": r.kind, "lines": r.lines, "yellow": r.yellow,
                    "unrecognised": sorted(set(r.unrecognised)), "missing": sorted(set(r.missing)),
                    "notes": r.notes} for r in results],
    } for name, _p, results, notes in items]
    return {"out_dir": out_dir, "report": report, "items": structured, "warning": WARNING}

def run(master_path, form_paths, out_root, confirm=True, assume_yes=False):
    res = fill_all(master_path, form_paths, out_root)
    print(res["report"])
    if confirm and not assume_yes:
        while input("\nType OK to confirm that you have read this: ").strip().upper() != "OK":
            pass
    return res["out_dir"]

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--master", required=True)
    ap.add_argument("--forms", nargs="+", required=True)
    ap.add_argument("--out", default="PortFill output")
    ap.add_argument("--yes", action="store_true", help="skip the confirmation prompt (for testing)")
    a = ap.parse_args()
    run(a.master, a.forms, a.out, assume_yes=a.yes)

if __name__ == "__main__":
    main()

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
    with open(os.path.join(out_dir, "Fill report.txt"), "w", encoding="utf-8") as f:
        f.write(report)
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

"""Quick self-test: German and Spanish crew lists with the test master (needs Test_Master.xlsx next to this file)."""
import os, sys, tempfile
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Side
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from master import Master
from forms import process_workbook

def make(path, headers):
    wb = Workbook(); ws = wb.active; ws.title = "Crew"
    thin = Side(style="thin")
    for j, h in enumerate(headers, 1):
        ws.cell(3, j, h).border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for r in range(4, 6):
        for j in range(1, len(headers) + 1):
            ws.cell(r, j).border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wb.save(path)

here = os.path.dirname(os.path.abspath(__file__))
master = Master(os.path.join(here, "Test_Master.xlsx"))
tmp = tempfile.mkdtemp()
cases = {
 "de": ["Nr.", "Nachname", "Vorname", "Dienstgrad", "Staatsangehörigkeit", "Geburtsdatum", "Geburtsort", "Geschlecht (M/F)"],
 "es": ["N.º", "Apellidos", "Nombre", "Rango", "Nacionalidad", "Fecha de nacimiento", "Lugar de nacimiento", "Sexo"],
 "en": ["No", "Family name", "Given names", "Rank or rating", "Nationality", "Date of birth", "Place of birth", "Gender"],
}
ok = True
for lang, headers in cases.items():
    p = os.path.join(tmp, f"crew_{lang}.xlsx")
    make(p, headers)
    out, results, notes = process_workbook(master, p, os.path.join(tmp, "out"))
    ws = load_workbook(out).active
    good = ws["B4"].value == "JANSEN" and ws["C4"].value == "Pieter Willem" and ws["D4"].value == "Master" and ws.cell(19, 2).value == "SILVA"
    print(lang, "OK" if good else "FAILED", "-", results[0].lines)
    ok &= good
print("ALL OK" if ok else "SOME TESTS FAILED")

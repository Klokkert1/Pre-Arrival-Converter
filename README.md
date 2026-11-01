# PortFill – test version

Fills port pre-arrival forms (Excel) from the ship's master Excel (`Master_FAL_Pre-arrival.xlsx`).
The original forms are never changed: a filled copy is saved in a new folder per port and voyage
(for example `Antwerp 2026-10-24`), together with a `Fill report.txt`.

## Start the window (needs Python 3.10+)

    pip install -r requirements.txt
    python portfill_app.py          (Windows: double-click start_windows.bat; Mac: start_mac.command)

The program runs a small page on your own computer (no internet needed) and opens it like an app window
(Edge or Chrome in app mode when available, otherwise your browser). Closing the window stops the program.

1. Choose the master file (type or paste the path, or use "Choose file").
2. Drop the port forms on the window (one workbook with tabs, or several files). It shows how many forms it recognises.
3. Choose the results folder (default: Documents/PortFill) and press "Fill forms".
4. Read the result list, then press "I have read this". The originals are never changed; a new folder per port and
   voyage (e.g. `Antwerp 2026-10-24`) is created. A second run never overwrites: it creates `... (2)`.

Command line (no window):

    python portfill.py --master Test_Master.xlsx --forms "ISPS form.xlsx" "NSW crew list.xlsx" --out output

## What it does

* Finds out by itself which sheet is which form, from the column headings and labels (English, German, Spanish).
* Fills crew list, passenger list, last 10 ports, ship's stores, waste, security information and ship name / IMO etc.
* Converts values to what the form asks: dates (format taken from the heading, e.g. `yyyy-mm-dd`, or from the cell),
  nationality to ISO-2 / ISO-3 code (from the dropdown list or the form's instructions), gender (`M/F` or `MALE/FEMALE`),
  ranks and document types to the form's dropdown lists, `Y/N`, names in capitals when the instructions ask for it.
* Adds extra rows when the form has too few lines (formatting copied; content below moves down; formulas follow).
* Leaves what it cannot fill empty and marks it **yellow**; values it could not convert safely are copied
  from the master and marked yellow, with a note in the report.

## Files

* `portfill_app.py` + `ui.html` – the window;  `portfill.py` – command line and report
* `forms.py` – recognising sheets and filling them
* `vocabulary.py` – words in English / German / Spanish (add more synonyms here, e.g. Dutch and Norwegian)
* `countries.py` – countries, nationalities, ISO codes
* `convert.py` – value conversions;  `matching.py` – label matching;  `master.py` – reads the master
* `make_test_master.py` – creates the test master with fake data;  `selftest.py` – quick self-test

## Known limits of this test version

* Two forms (NSW Security, NSW Waste) store some dropdown lists in an Excel extension format that cannot be kept
  when the copy is saved. The values are filled in, the dropdown arrows in those cells are lost (the report says so).
* Form controls (checkboxes/buttons) are not kept. Arrival/Departure tick boxes are left to the user (yellow).
* Sheets Certificates, Hailing, Maritime Declaration of Health (apart from ship data and port list) are not filled yet.
* Stores: quantities are only added up when the units can be converted safely (litres, kg, pieces).
  'bottles' or 'cartons' are not converted; the cell stays empty and yellow.
* Not tested yet on real Windows / Mac: opening as an app window, and the 'Choose file' / 'Change folder' dialogs
  (if a dialog does not open, type the path in the box). Tested here in a headless browser.
* Not yet built: licence check, Windows .exe / Mac .app, Dutch and Norwegian words.
* Windows .exe later: `pip install pyinstaller` then `pyinstaller --onefile portfill.py`.

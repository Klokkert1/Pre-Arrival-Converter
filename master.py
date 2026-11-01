"""Reads the master Excel (the Master_FAL_Pre-arrival.xlsx template)."""
import datetime
from openpyxl import load_workbook
from matching import match, norm
from vocabulary import PERSON, PORTS, WASTE

# label in master kv-sheets -> field id (same ids as vocabulary.FIELDS)
LABEL_TO_ID = {
    "Name of ship": "ship_name", "IMO number": "imo", "Call sign": "call_sign", "Flag State of ship": "flag",
    "Name of master": "master_name", "Type of ship": "ship_type", "Gross tonnage": "gt", "Net tonnage": "nt",
    "Voyage number": "voyage_no", "Arrival / Departure": "movement", "Port of arrival/departure": "port",
    "Date and time of arrival/departure": "eta", "Last port of call": "last_port", "Next port of call": "next_port",
    "Name and contact details of ship’s agent": "agent",
    "Valid ISSC": "issc_valid", "ISSC issuing company / authority": "issc_issuer", "ISSC issuer type": "issc_issuer_type",
    "SSP on board": "ssp_onboard", "ISSC expiry date": "issc_expiry", "CSO given name": "cso_given",
    "CSO family name": "cso_family", "CSO phone 24/7": "cso_phone", "CSO e-mail": "cso_email",
    "Company name": "company_name", "Company IMO number": "company_imo", "Ship configuration": "ship_config",
    "Condition of cargo tanks": "cargo_condition", "Ballast tanks": "ballast",
}

class Master:
    def __init__(self, path):
        self.path = path
        wb = load_workbook(path, data_only=True)
        self.wb = wb
        self.fields = {}
        for sheet in ("Ship", "Voyage", "Security"):
            if sheet not in wb.sheetnames:
                continue
            ws = wb[sheet]
            for r in range(4, ws.max_row + 1):
                label, val = ws.cell(r, 2).value, ws.cell(r, 3).value
                if label in LABEL_TO_ID and val not in (None, ""):
                    self.fields[LABEL_TO_ID[label]] = val
        self.crew = self._table("Crew_List", PERSON, key="family_name")
        self.pax = self._table("Passengers", PERSON, key="family_name")
        self.ports = self._table("Last_10_Ports", PORTS, key="port", header_row=5, first=6, last=15)
        self.stores = self._stores()
        self.waste = self._waste()

    def _table(self, sheet, vocab, key, header_row=5, first=6, last=None):
        if sheet not in self.wb.sheetnames:
            return []
        ws = self.wb[sheet]
        cols = {}
        for c in range(1, ws.max_column + 1):
            h = ws.cell(header_row, c).value
            if h:
                concept, _ = match(h, vocab)
                if concept and concept not in cols.values():
                    cols[c] = concept
        rows = []
        for r in range(first, (last or ws.max_row) + 1):
            rec = {cols[c]: ws.cell(r, c).value for c in cols}
            if rec.get(key) in (None, ""):
                continue
            rec = {k: v for k, v in rec.items() if v not in (None, "")}
            rows.append(rec)
        return rows

    def _stores(self):
        if "Ship_Stores" not in self.wb.sheetnames:
            return []
        ws = self.wb["Ship_Stores"]
        out = []
        for r in range(6, 66):
            name = ws.cell(r, 2).value
            if name in (None, ""):
                continue
            out.append({"category": ws.cell(r, 1).value or "", "name": str(name), "qty": ws.cell(r, 3).value,
                        "unit": ws.cell(r, 4).value or "", "place": ws.cell(r, 5).value})
        return out

    def _waste(self):
        if "Waste" not in self.wb.sheetnames:
            return {"header": {}, "rows": []}
        ws = self.wb["Waste"]
        header = {"w_last_port": ws["C4"].value, "w_last_date": ws["C5"].value, "w_next_port": ws["C6"].value}
        rows = []
        for r in range(9, ws.max_row + 1):
            name = ws.cell(r, 2).value
            if name in (None, ""):
                continue
            vals = {"w_deliver": ws.cell(r, 3).value, "w_capacity": ws.cell(r, 4).value,
                    "w_onboard": ws.cell(r, 5).value, "w_generated": ws.cell(r, 6).value,
                    "w_delivered_last": ws.cell(r, 7).value, "w_remarks": ws.cell(r, 8).value}
            vals = {k: v for k, v in vals.items() if v not in (None, "")}
            rows.append({"code": str(ws.cell(r, 1).value or "").strip(), "name": str(name), "values": vals})
        return {"header": header, "rows": rows}

    def folder_name(self):
        port = str(self.fields.get("port", "Port")).strip() or "Port"
        d = self.fields.get("eta")
        stamp = d.strftime("%Y-%m-%d") if isinstance(d, (datetime.date, datetime.datetime)) else ""
        return f"{port} {stamp}".strip()

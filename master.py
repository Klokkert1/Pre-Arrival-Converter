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
    "Total number of ballast tanks": "bw_tanks", "Number of tanks in ballast": "bw_in_ballast",
    "Number of tanks exchanged": "bw_exchanged", "Number of tanks not exchanged": "bw_not_exchanged",
    "Total ballast water on board (m3)": "bw_on_board", "Total ballast water capacity (m3)": "bw_capacity",
    "Ballast water volume at arrival (m3)": "bw_vol_arrival", "Ballast water volume at departure (m3)": "bw_vol_departure",
    "Ballast water management plan on board?": "bw_plan", "Ballast water treatment system on board?": "bw_treatment",
    "Type of ballast water treatment": "bw_treatment_type", "Ballast water exchange in port?": "bw_exchange_port",
    "Ballast water discharge in port?": "bw_discharge", "Volume of ballast water to be discharged (m3)": "bw_discharge_vol",
    "Ballast water loaded in port?": "bw_loaded", "Volume of ballast water to be loaded (m3)": "bw_loaded_vol",
    "IMO ballast water guidelines on board?": "bw_guidelines",
    # added from the company's own forms
    "Official number": "official_no", "Port of registry": "port_registry", "Inmarsat call numbers": "inmarsat",
    "UN/LOCODE of port of arrival/departure": "port_locode", "Date and time of arrival (ETA / ATA)": "eta_arrival",
    "Date and time of departure (ETD / ATD)": "etd_departure", "Period of stay": "period_of_stay",
    "UN/LOCODE of last port of call": "last_port_locode", "UN/LOCODE of next port of call": "next_port_locode",
    "Port facility of arrival (if known)": "port_facility", "Primary purpose of call": "call_purpose",
    "Security level the ship is currently operating at": "ship_sec_level",
    "Location of ship at time of report": "ship_location",
    "Dangerous substances as cargo (ISPS classes)?": "dg_cargo",
    "Security procedures maintained during ship-to-ship activities?": "sts_procedures",
    "Security-related matter to report?": "sec_matter", "Details of security-related matter": "sec_matter_details",
    "Title or position of person providing the information": "reporter_title",
    "Date / time / place of completion of report": "report_completed",
    "Terminal / berth": "berth", "ISSC certificate number": "issc_number", "ISSC date of issue": "issc_issued",
    "Sanitation control: certificate or exemption": "san_type", "Sanitation certificate / exemption issued at": "san_place",
    "Sanitation certificate / exemption dated": "san_date", "Re-inspection required?": "reinspection",
    "Deadweight (DWT, tonnes)": "dwt", "Length overall (LOA, m)": "length", "Beam (m)": "breadth",
    "Draught forward (m)": "draft_fore", "Draught aft (m)": "draft_aft", "MMSI": "mmsi",
    "Ship e-mail": "ship_email", "Ship phone / satcom": "ship_phone", "Owner / manager": "owner",
    "Certificate of registry (port; date; number)": "cert_registry", "Brief particulars of voyage": "voyage_particulars",
    "Brief description of the cargo": "cargo_desc", "Number of crew": "n_crew", "Number of passengers": "n_pax",
    "Remarks": "remarks", "Cargo declaration": "att_cargo", "Ship’s stores declaration": "att_stores",
    "Crew list": "att_crew", "Passenger list": "att_pax", "Crew’s effects declaration (arrival only)": "att_effects",
    "Passengers": "nil_passengers", "Animals": "nil_animals", "Arms and ammunition": "nil_arms",
    "Stowaways": "nil_stowaways",
}

# Maritime Declaration of Health: (concept id, label on sheet Health, [(detail id, detail label)])
HEALTH_QUESTIONS = [
    ("death", "Death on board during the voyage (other than by accident)?", [("death_count", "Total number of deaths")]),
    ("infectious", "Case or suspected case of infectious disease on board during the voyage?", []),
    ("ill_pax", "Total number of ill passengers greater than normal / expected?", [("ill_count", "How many ill persons")]),
    ("ill_now", "Any ill person on board now?", []),
    ("doctor", "Medical practitioner consulted?", []),
    ("condition", "Condition on board which may lead to infection or spread of disease?", []),
    ("sanitary", "Sanitary measure (quarantine, isolation, disinfection, decontamination) applied on board?",
     [("san_measure", "Sanitary measure – type, place and date")]),
    ("stowaways", "Stowaways found on board?", [("stow_where", "Stowaways – where did they join the ship (if known)?")]),
    ("animal", "Sick animal or pet on board?", []),
    ("affected", "Ship has visited an area affected as identified by the WHO?",
     [("affected_where", "Affected area – port and date of visit")]),
    ("rats", "Plague suspected among rats or mice, or abnormal mortality among them?", []),
    ("doctor_onboard", "Doctor on board?", []),
]
HEALTH_LABELS = {}
for _c, _l, _d in HEALTH_QUESTIONS:
    HEALTH_LABELS[_l] = _c
    for _dc, _dl in _d:
        HEALTH_LABELS[_dl] = _dc
HEALTH_LABELS["Particulars of cases / deaths / sick persons (schedule)"] = "particulars"


class Master:
    def __init__(self, path):
        self.path = path
        wb = load_workbook(path, data_only=True)
        self.wb = wb
        self.fields = {}
        self.health = {}
        for sheet in ("Ship", "Voyage", "Security", "Health", "Nil_List", "Ballast"):
            if sheet not in wb.sheetnames:
                continue
            ws = wb[sheet]
            for r in range(4, ws.max_row + 1):
                label, val = ws.cell(r, 2).value, ws.cell(r, 3).value
                if label in LABEL_TO_ID and val not in (None, ""):
                    self.fields[LABEL_TO_ID[label]] = val
                elif sheet == "Health" and label in HEALTH_LABELS and val not in (None, ""):
                    self.health[HEALTH_LABELS[label]] = val
        self.crew = self._table("Crew_List", PERSON, key="family_name")
        self._crew_documents()
        self._crew_effects()
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

    def _crew_documents(self):
        """New master: passport and seaman's book in separate columns. Forms that ask for ONE identity document
        get the passport (or the seaman's book when there is no passport); both stay available as pp_* / sb_*."""
        if "Crew_List" not in self.wb.sheetnames:
            return
        ws = self.wb["Crew_List"]
        cols = {}
        for c in range(1, ws.max_column + 1):
            h = str(ws.cell(5, c).value or "").strip().lower()
            for pre, tag in (("passport", "pp"), ("seaman", "sb")):
                if h.startswith(pre):
                    kind = "no" if h.endswith(" no.") else "state" if "state" in h else "exp" if "expiry" in h else None
                    if kind:
                        cols[c] = f"{tag}_{kind}"
        if not cols:
            return
        recs = iter(self.crew)
        for r in range(6, ws.max_row + 1):
            if ws.cell(r, 2).value in (None, ""):
                continue
            rec = next(recs, None)
            if rec is None:
                break
            for k in ("id_type", "id_number", "id_issue_country", "id_expiry"):
                rec.pop(k, None)
            for c, k in cols.items():
                v = ws.cell(r, c).value
                if v not in (None, ""):
                    rec[k] = v
            for tag, name in (("pp", "Passport"), ("sb", "Seaman's book")):
                if rec.get(tag + "_no"):
                    rec.update({"id_type": name, "id_number": rec[tag + "_no"],
                                "id_issue_country": rec.get(tag + "_state"), "id_expiry": rec.get(tag + "_exp")})
                    for k in [k for k, v in rec.items() if v is None]:
                        del rec[k]
                    break

    def _crew_effects(self):
        """Crew_Effects sheet (FAL 4): cigarettes / tobacco / wines-spirits / cash / electronics per crew member."""
        if "Crew_Effects" not in self.wb.sheetnames or "Crew_List" not in self.wb.sheetnames:
            return
        ce, cl = self.wb["Crew_Effects"], self.wb["Crew_List"]
        keys = {5: "eff_cig", 6: "eff_tob", 7: "eff_wine", 8: "eff_cash", 9: "eff_elec"}
        recs = iter(self.crew)
        for r in range(6, cl.max_row + 1):
            if cl.cell(r, 2).value in (None, ""):
                continue
            rec = next(recs, None)
            if rec is None:
                break
            for c, k in keys.items():
                v = ce.cell(r, c).value
                if v not in (None, ""):
                    rec[k] = v

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

"""Vocabulary: what form labels / column headers mean, in English, German and Spanish.

Each concept maps to a list of synonyms (written in plain lowercase, no punctuation).
- A one-word synonym only matches when the whole label equals it.
- A multi-word synonym matches when all its words occur in the label.
To add a language or a new wording: just add more synonyms. Nothing else needs to change.
"""

# ---------- Table columns: crew list and passenger list ----------
PERSON = {
    "serial": ["no", "nr", "number", "serial no", "serial number", "lfd nr", "nummer", "numero", "num"],
    "effects_text": ["effects ineligible for relief", "effects ineligible", "crew effects", "crews effects", "effects",
                     "personal effects", "effekten"],
    "family_name": ["family name", "surname", "last name", "lastname", "nachname", "familienname", "zuname",
                    "apellidos", "apellido", "primer apellido"],
    "given_names": ["given names", "given name", "first name", "forename", "forenames", "vorname", "vornamen",
                    "nombres", "nombre de pila", "nombre"],
    "rank": ["rank or rating", "rank rating", "rank", "duty of crew", "function", "dienstgrad", "dienstrang",
             "funktion", "rango", "cargo", "puesto"],
    "row_ship_name": ["shipname", "ship name", "vessel name", "name of ship", "name of vessel"],
    "second_name": ["second name", "secondname", "middle name", "middle names", "second given name"],
    "nationality": ["nationality", "staatsangehorigkeit", "nationalitat", "nacionalidad"],
    "dob": ["date of birth", "birth date", "geburtsdatum", "fecha de nacimiento", "dob"],
    "pob": ["place of birth", "birth place", "geburtsort", "lugar de nacimiento"],
    "gender": ["gender", "sex", "geschlecht", "sexo", "genero"],
    "id_type": ["nature of identity document", "type of identity document", "identity document type",
                "nature of id document", "type of id document", "art des ausweises", "ausweisart",
                "tipo de documento", "tipo de documento de identidad"],
    "id_number": ["number of identity document", "identity document number", "identity document no",
                  "number of id document", "id document number", "passport number", "ausweisnummer",
                  "numero de documento", "numero del documento de identidad"],
    "id_issue_date": ["date of issue", "issue date", "ausstellungsdatum", "fecha de expedicion"],
    "id_expiry": ["date of validity", "expiry date", "expire date", "expiration date", "valid until", "date of expiry",
                  "gultig bis", "ablaufdatum", "fecha de caducidad", "fecha de vencimiento"],
    "id_issue_country": ["country of issue", "issuing state", "issuing country", "ausstellungsland",
                         "ausstellender staat", "pais de expedicion"],
    "issuer_entity": ["name of issuer entity", "issuer entity", "issuing authority", "name of issuer",
                      "ausstellende behorde", "autoridad expedidora"],
    "visa_no": ["visa residence permit number", "serial number of visa", "visa residence permit no",
                "visa number", "visumnummer", "numero de visado"],
    "visa_type": ["type of visa", "visa type", "visumart", "tipo de visado"],
    "crew_effects": ["crew effects", "effects ineligible", "crew s effects", "effekten", "efectos"],
    "port_emb": ["port of embarkation", "embarkation port", "einschiffungshafen", "puerto de embarque"],
    "port_disemb": ["port of disembarkation", "disembarkation port", "port of desembarkation",
                    "ausschiffungshafen", "puerto de desembarque"],
    "date_emb": ["date of embarkation", "embarkation date", "einschiffungsdatum", "fecha de embarque"],
    "date_disemb": ["date of disembarkation", "disembarkation date", "date of desembarkation",
                    "ausschiffungsdatum", "fecha de desembarque"],
    "transit": ["transit", "transit passenger", "license to be ashore", "shore leave", "landgang",
                "en transito", "stay on board"],
    "remarks": ["observations remarks", "observations", "remarks", "notes", "bemerkungen", "observaciones"],
}

# ---------- Table columns: port calls ----------
PORTS = {
    "serial": ["no", "nr", "number"],
    "port": ["port", "port or un locode", "port locode", "port name", "port of call", "hafen", "puerto"],
    "locode": ["un locode", "locode", "unlocode", "locode of port"],
    "country": ["country", "country where port is located", "land", "pais"],
    "arrival": ["date of arrival", "arrival date", "ata", "date from", "ankunft", "fecha de llegada", "llegada"],
    "departure": ["date of departure", "departure date", "atd", "date to", "abfahrt", "fecha de salida", "salida"],
    "facility": ["port facility", "facility gisis", "port facility no", "hafenanlage", "instalacion portuaria"],
    "sec_level": ["security level", "sicherheitsstufe", "nivel de proteccion", "nivel de seguridad"],
    "measures": ["special or additional security", "additional security measures", "zusatzliche sicherheitsmassnahmen",
                 "medidas de proteccion adicionales"],
    "next_port": ["next port", "next port if known", "nachster hafen", "proximo puerto"],
    "transit_time": ["transit time", "fahrzeit"],
    "eta_next": ["eta next port"],
}

# ---------- Table columns: waste ----------
WASTE = {
    "w_type": ["waste type", "type of waste", "abfallart", "tipo de residuo"],
    "w_deliver": ["to be delivered", "quantity to be delivered", "qtty to be delivered", "zur abgabe",
                  "a entregar"],
    "w_capacity": ["storage capacity", "max storage capacity", "lagerkapazitat", "capacidad de almacenamiento"],
    "w_onboard": ["kept on board", "waste on board", "on board", "an bord", "a bordo"],
    "w_generated": ["to be generated", "estimated amount", "estimated quantity", "generated",
                    "voraussichtlich", "estimada"],
    "w_delivered_last": ["delivered at last port", "delivered at the last port", "delivered at last port of delivery"],
    "w_next_port": ["next port of delivery", "port of delivery of remaining waste", "nachster abgabehafen",
                    "proximo puerto de entrega"],
    "w_remarks": ["remarks"],
}

# ---------- Single fields (label / value forms) ----------
FIELDS = {
    "ship_name": ["name of ship", "vessel name", "ship name", "name of vessel", "vessel", "ship", "schiffsname",
                  "name des schiffes", "nombre del buque", "nombre del barco", "buque"],
    "imo": ["imo number", "imo no", "imo", "imo nummer", "numero omi", "numero imo"],
    "call_sign": ["call sign", "callsign", "rufzeichen", "indicativo de llamada"],
    "flag": ["flag state", "flag", "flagge", "flaggenstaat", "bandera"],
    "master_name": ["name of master", "master", "kapitan", "capitan"],
    "gt": ["gross tonnage", "gt", "bruttoraumzahl", "arqueo bruto"],
    "nt": ["net tonnage", "nt", "nettoraumzahl", "arqueo neto"],
    "ship_type": ["type of ship", "ship type", "vessel type", "type of vessel", "schiffstyp", "tipo de buque"],
    "portnet_no": ["portnet number", "portnet no"],
    "voyage_no": ["voyage number", "voyage no", "reisenummer", "numero de viaje"],
    "port_mv": ["port of arrival departure", "port of arrival or departure"],
    "port_arr": ["port of arrival", "arrival port", "ankunftshafen"],
    "port_dep": ["port of departure", "departure port", "abfahrtshafen"],
    "date_mv": ["date of arrival departure", "date of arrival or departure"],
    "date_arr": ["date of arrival", "arrival date", "ankunftsdatum"],
    "date_dep": ["date of departure", "departure date", "abfahrtsdatum"],
    "last_port": ["last port of call", "last port", "letzter anlaufhafen"],
    "eta_dt": ["expected date and time of arrival", "date and time of arrival", "eta", "eta time", "eta date and time", "estimated time of arrival", "estimated date and time of arrival"],
    "port_registry": ["port of registry", "registry port", "heimathafen"],
    "length": ["length overall", "loa", "length over all", "length of ship", "length of vessel", "ship length", "vessel length"],
    "breadth": ["beam", "breadth", "moulded breadth", "breadth moulded", "breadth of ship", "breadth of vessel", "beam of vessel"],
    "dwt": ["deadweight", "dwt", "summer dwt", "summer deadweight", "deadweight tonnage"],
    "draft_fa": ["draft fwd aft", "draught fwd aft", "drafts fwd aft", "draft forward aft", "draught forward aft"],
    "draft_fore": ["draft forward", "forward draft", "draught forward", "fwd draft", "draft fwd", "draught fwd"],
    "draft_aft": ["draft aft", "aft draft", "draught aft"],
    "etd_dt": ["etd", "etd time", "etd date and time", "estimated time of departure", "estimated date and time of departure"],
    "berth": ["berth", "berth bollard nos", "berth number", "berth name", "berth no", "berth bollard"],
    "agent": ["vessel agent", "ship s agent", "ships agent", "agent", "name of agent", "agent s name", "name of ship s agent"],
    "owner": ["owner", "owners", "ship owner", "registered owner", "owner manager", "vessel owner"],
    "ship_phone": ["vessel telephone", "ship telephone", "ship phone", "vessel phone", "telephone number of ship"],
    "ship_email": ["vessel e mail", "ship e mail", "vessel email", "ship email", "e mail of ship", "e mail of vessel"],
    "mmsi": ["mmsi", "mmsi number", "mmsi no"],
    "official_no": ["official number", "official no", "official number of ship"],
    "movement_arrival": ["arrival"],
    "movement_departure": ["departure"],
    # security
    "issc_valid": ["valid issc"],
    "issc_issuer": ["issc issuing company", "issc issuer"],
    "issc_issuer_type": ["issc type"],
    "ssp_onboard": ["ssp onboard", "ssp on board"],
    "issc_expiry": ["expiry date"],
    "cso_given": ["cso given name", "cso first name"],
    "cso_family": ["cso family name", "cso surname"],
    "cso_phone": ["phone 24 7", "cso phone", "phone"],
    "cso_email": ["email", "e mail", "cso email"],
    "company_name": ["company name"],
    "ship_config": ["ship configuration"],
    "cargo_condition": ["condition cargo", "condition of cargo tanks"],
    "ballast": ["ballasttanks", "ballast tanks"],
    # waste header fields
    "w_last_port": ["last port delivered", "last port of delivery"],
    "w_last_date": ["date of delivery"],
    "w_next_port_hdr": ["next port of delivery"],
}
SECURITY_FIELDS = {"issc_valid", "issc_issuer", "issc_issuer_type", "ssp_onboard", "issc_expiry", "cso_given",
                   "cso_family", "cso_phone", "cso_email", "company_name", "ship_config", "cargo_condition", "ballast"}
HEADER_FIELDS = {"length", "breadth", "dwt", "draft_fa", "draft_fore", "draft_aft", "etd_dt", "berth", "agent", "owner", "ship_phone", "ship_email", "mmsi", "official_no", "eta_dt", "port_registry", "nt", "voyage_no", "port_mv", "port_arr", "port_dep", "date_mv", "date_arr", "date_dep", "last_port", "ship_name", "imo", "portnet_no", "movement_arrival", "movement_departure", "flag", "call_sign",
                 "gt", "master_name", "ship_type"}
_PX = {"date", "time", "facility", "expected", "ship", "berth", "agent", "estimated"}
FIELD_EXCLUDE = {"port_mv": _PX, "port_arr": _PX, "port_dep": _PX, "last_port": {"facility", "date", "time", "delivery"},
                 "date_mv": {"time", "port"}, "date_arr": {"time", "port"}, "date_dep": {"time", "port"},
                 "voyage_no": {"particulars"},
                 "ship_name": {"agent", "master", "owner", "manager", "company", "officer"},
                 "flag": {"agent"}}
WASTE_HEADER_FIELDS = {"w_last_port", "w_last_date", "w_next_port_hdr"}
# Concepts that the program knows about but where it is not clear how the form wants them: left empty + yellow.
UNSURE = {"movement_arrival", "movement_departure"}
# Concepts the master does not have at all.
NOT_IN_MASTER = {"portnet_no", "id_issue_date", "issuer_entity", "crew_effects", "port_emb", "port_disemb",
                 "date_emb", "date_disemb", "visa_type", "transit_time", "eta_next"}

# ---------- Ship's stores: classes of articles ----------
# (class, keywords in the form label or in the master item name). First match with the longest keyword wins.
STORE_CLASSES = [
    ("other_alcohol", ["other spirits", "other alcohol", "andere spirituosen", "otros licores"]),
    ("spirits", ["spirits", "whisky", "whiskey", "vodka", "gin", "rum", "brandy", "cognac", "liquor", "liqueur",
                 "spirituosen", "schnaps", "licores", "aguardiente"]),
    ("beer", ["beer", "bier", "cerveza", "lager"]),
    ("wine", ["wine", "wein", "vino"]),
    ("cigarettes", ["cigarettes", "cigarette", "zigaretten", "cigarrillos"]),
    ("cigars", ["cigars", "cigar", "zigarren", "puros"]),
    ("tobacco", ["tobacco", "tabacco", "tabak", "tabaco"]),
    ("mgo", ["mgo", "mdo", "gas oil", "diesel", "marine gas oil", "gasoil"]),
    ("fuel_oil", ["fuel oil", "hfo", "ifo", "heavy fuel", "bunker", "schweroel", "fuelol", "fueloleo"]),
    ("lub_oil", ["lub oil", "lube oil", "lubricating oil", "lubrication oil", "lubrication", "lubricant", "schmierol", "aceite lubricante"]),
    ("medicines", ["medicines", "medicine", "medikamente", "medicamentos"]),
    ("meat", ["meat", "fleisch", "carne"]),
    ("firearms", ["fire arms", "firearms", "ammunition", "waffen", "armas"]),
    ("coffee", ["coffee", "kaffee", "cafe"]),
    ("fresh_water", ["fresh water", "freshwater", "drinking water", "potable water", "trinkwasser", "agua dulce"]),
    ("paint", ["paint", "thinner", "farbe", "pintura"]),
    ("chemicals", ["chemicals", "chemical", "chemikalien", "quimicos"]),
    ("pyrotechnics", ["pyrotechnics", "pyrotechnic", "flares", "parachute rockets", "distress rockets", "signal rockets"]),
    ("line_thrower", ["line thrower", "line throwing", "rocket line"]),
    ("provisions", ["provisions", "provision", "food stores", "proviant", "provisiones"]),
]


# ---------- Extra word lists (one text file per language, see the folder 'words') ----------
def _load_words():
    """Add the synonyms from words/*.txt (bundled) and from ~/.portfill/words/*.txt (your own additions).
    File format:  [section] on its own line, then   concept = word one, word two, ...   (# starts a comment)."""
    import os, sys, glob
    from matching import norm
    here = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    home = os.environ.get("PORTFILL_HOME") or os.path.join(os.path.expanduser("~"), ".portfill")
    files = sorted(glob.glob(os.path.join(here, "words", "*.txt"))) + sorted(glob.glob(os.path.join(home, "words", "*.txt")))
    targets = {"person": PERSON, "ports": PORTS, "waste": WASTE, "fields": FIELDS}
    extra_stores = {"article": [], "quantity": [], "unit": [], "location": []}
    store_cls = dict(STORE_CLASSES_MAP)
    problems = []
    for fn in files:
        section = None
        try:
            lines = open(fn, encoding="utf-8-sig").read().splitlines()
        except OSError:
            continue
        for n, line in enumerate(lines, 1):
            line = line.split("#", 1)[0].strip()
            if not line:
                continue
            if line.startswith("[") and line.endswith("]"):
                section = line[1:-1].strip().lower()
                continue
            if "=" not in line or section is None:
                problems.append(f"{os.path.basename(fn)} line {n}: not understood")
                continue
            key, _, vals = line.partition("=")
            key = key.strip().lower()
            words = [norm(w) for w in vals.split(",") if norm(w)]
            if section in targets:
                if key not in targets[section]:
                    problems.append(f"{os.path.basename(fn)} line {n}: unknown field '{key}' in [{section}]")
                    continue
                for w in words:
                    if w not in targets[section][key]:
                        targets[section][key].append(w)
            elif section == "store_classes":
                if key not in store_cls:
                    problems.append(f"{os.path.basename(fn)} line {n}: unknown article class '{key}'")
                    continue
                for w in words:
                    if w not in store_cls[key]:
                        store_cls[key].append(w)
            elif section == "stores" and key in extra_stores:
                extra_stores[key] += [w for w in words if w not in extra_stores[key]]
            else:
                problems.append(f"{os.path.basename(fn)} line {n}: unknown section/field")
    return extra_stores, problems


STORE_CLASSES_MAP = {name: kws for name, kws in STORE_CLASSES}
EXTRA_STORES, WORD_PROBLEMS = _load_words()

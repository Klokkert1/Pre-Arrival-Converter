"""Vocabulary: what form labels / column headers mean, in English, German and Spanish.

Each concept maps to a list of synonyms (written in plain lowercase, no punctuation).
- A one-word synonym only matches when the whole label equals it.
- A multi-word synonym matches when all its words occur in the label.
To add a language or a new wording: just add more synonyms. Nothing else needs to change.
"""

# ---------- Table columns: crew list and passenger list ----------
PERSON = {
    "serial": ["no", "nr", "number", "serial no", "serial number", "lfd nr", "nummer", "numero", "num"],
    "family_name": ["family name", "surname", "last name", "lastname", "nachname", "familienname", "zuname",
                    "apellidos", "apellido", "primer apellido"],
    "given_names": ["given names", "given name", "first name", "forename", "forenames", "vorname", "vornamen",
                    "nombres", "nombre de pila", "nombre"],
    "rank": ["rank or rating", "rank rating", "rank", "duty of crew", "function", "dienstgrad", "dienstrang",
             "funktion", "rango", "cargo", "puesto"],
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
    "id_expiry": ["date of validity", "expiry date", "expiration date", "valid until", "date of expiry",
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
    "locode": ["un locode", "locode", "unlocode"],
    "country": ["country", "land", "pais"],
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
HEADER_FIELDS = {"ship_name", "imo", "portnet_no", "movement_arrival", "movement_departure", "flag", "call_sign",
                 "gt", "master_name", "ship_type"}
FIELD_EXCLUDE = {"ship_name": {"agent", "master", "owner", "manager", "company", "officer"},
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

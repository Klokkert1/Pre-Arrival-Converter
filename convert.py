"""Converting master values into the form a port form asks for."""
import re, datetime, difflib
from matching import norm, compact, hint_text
import countries

RANK_ALIASES = {
    "chief officer": "ChiefOfficer", "chief mate": "ChiefMate", "master": "Master", "captain": "Captain",
    "first officer": "FirstOfficer", "second officer": "SecondOfficer", "third officer": "ThirdOfficer",
    "second mate": "SecondMate", "third mate": "ThirdMate", "chief engineer": "ChiefEngineer",
    "second engineer": "SecondAsstEngineer", "third engineer": "ThirdAsstEngineer",
    "fourth engineer": "FourthAsstEngineer", "first engineer": "FirstEngineer",
    "electrician": "Electrician", "eto": "ElectricalOfficer", "electro technical officer": "ElectricalOfficer",
    "bosun": "Bosun", "boatswain": "Bosun", "able seaman": "AbleSeaman", "ab": "AbleSeaman",
    "ordinary seaman": "OrdinarySeaman", "os": "OrdinarySeaman", "oiler": "Oiler", "motorman": "Motorman",
    "cook": "Cook", "chief cook": "ChiefCook", "steward": "Steward", "deck cadet": "Cadet", "cadet": "Cadet",
    "engine cadet": "EngineerCadet", "wiper": "Wiper", "fitter": "Fitter", "pumpman": "Pumpman",
    "radio officer": "RadioOfficer", "messman": "Messman", "deckhand": "Deckhand",
}

def to_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if v in (None, ""):
        return None
    s = str(v).strip()
    for f in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%d.%m.%Y", "%d-%m-%Y %H:%M", "%Y-%m-%d %H:%M:%S", "%d-%m-%y"):
        try:
            return datetime.datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None

def is_date_format(fmt):
    f = (fmt or "").lower()
    return f != "general" and bool(re.search(r"d{1,4}|y{2,4}", f)) and bool(re.search(r"m{1,4}", f))

def date_hint_format(hints):
    m = re.search(r"(yyyy|dd|mm)([-/. ])(yyyy|dd|mm)\2(yyyy|dd|mm)", hints.lower())
    return m.group(0) if m else None

def choose_from_list(value, allowed, aliases=None):
    """Pick the allowed list entry that matches a value. Returns entry or None."""
    if value in (None, "") or not allowed:
        return None
    comp = {compact(a): a for a in allowed if a not in (None, "")}
    c = compact(value)
    if c in comp:
        return comp[c]
    if aliases:
        a = aliases.get(norm(value))
        if a and compact(a) in comp:
            return comp[compact(a)]
    close = difflib.get_close_matches(c, list(comp), n=1, cutoff=0.88)
    return comp[close[0]] if close else None

def convert_country(value, allowed, label_text):
    """Nationality or country -> what the form wants (ISO-2, ISO-3 or name). Returns (value, ok)."""
    if value in (None, ""):
        return None, True
    rec = countries.lookup(value)
    if allowed:
        sample = [str(a) for a in allowed if a not in (None, "")][:40]
        if sample and all(len(a) == 2 and a.isupper() for a in sample):
            kind = "a2"
        elif sample and all(len(a) == 3 and a.isupper() for a in sample):
            kind = "a3"
        else:
            kind = "name"
        if rec:
            if kind in ("a2", "a3"):
                cand = rec[kind]
            else:
                cand = choose_from_list(rec["name"], allowed) or choose_from_list(value, allowed)
                if cand is None:
                    for other in (rec["name"],):
                        cand = choose_from_list(other, allowed)
            if cand and choose_from_list(cand, allowed):
                return choose_from_list(cand, allowed), True
        return None, False
    t = norm(label_text)
    if rec:
        if "iso" in t and ("3" in t or "alpha 3" in t):
            return rec["a3"], True
        if "iso" in t or "code" in t:
            return rec["a2"], True
    return value, True  # no restriction in the form: keep the master text

def convert_gender(value, hints, allowed):
    if value in (None, ""):
        return None, True
    v = norm(value)
    male = v in ("m", "male", "man", "mann", "masculino", "hombre")
    female = v in ("f", "female", "woman", "frau", "femenino", "mujer", "w")
    if not (male or female):
        return None, False
    if allowed:
        pick = choose_from_list("male" if male else "female", allowed)
        if pick:
            return pick, True
    h = hint_text(hints) or hints
    toks = {t for t in re.split(r"[^A-Za-z]+", h) if t}
    up = {t.upper() for t in toks}
    if "MALE" in up or "FEMALE" in up:
        word = "male" if male else "female"
        return (word.upper() if any(t.isupper() for t in toks if len(t) > 1) else word.capitalize()), True
    if "M" in up or "F" in up:
        return ("M" if male else "F"), True
    return ("M" if male else "F"), True

def convert_yesno(value, hints, allowed):
    if value in (None, ""):
        return None, True
    v = norm(value)
    yes = v in ("y", "yes", "ja", "si", "true", "1")
    no = v in ("n", "no", "nein", "false", "0")
    if not (yes or no):
        return None, False
    if allowed:
        pick = choose_from_list("yes" if yes else "no", allowed)
        if pick:
            return pick, True
    h = hint_text(hints)
    if re.search(r"\by\b", h.lower()):
        return ("Y" if yes else "N"), True
    if "ja" in h.lower():
        return ("Ja" if yes else "Nein"), True
    return ("Yes" if yes else "No"), True

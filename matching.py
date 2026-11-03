"""Text normalisation and label matching."""
import re, unicodedata

STOP_SPLIT = re.compile(r"[^a-z0-9 ]+")

_TRANSLIT = str.maketrans({"ß": "ss", "ø": "o", "Ø": "O", "æ": "ae", "Æ": "AE", "œ": "oe", "Œ": "OE", "ł": "l",
                           "Ł": "L", "đ": "d", "Đ": "D", "ı": "i", "þ": "th"})


def norm(text):
    t = unicodedata.normalize("NFKD", str(text).translate(_TRANSLIT))
    t = "".join(c for c in t if not unicodedata.combining(c)).lower().replace("_", " ")
    t = STOP_SPLIT.sub(" ", t)
    return re.sub(r"\s+", " ", t).strip()

def compact(text):
    return re.sub(r"[^a-z0-9]", "", norm(text))

def variants(label):
    s = str(label)
    no_paren = re.sub(r"\([^)]*\)", " ", s)
    nonum = re.sub(r"^\s*\d+(\.\d+)*\.?\s+(?=\D)", "", no_paren)
    return [v for v in {norm(no_paren), norm(s), norm(nonum)} if v]

def hint_text(label):
    """Text in parentheses, e.g. '(yyyy-mm-dd)' or '(M, F)'."""
    return " ".join(re.findall(r"\(([^)]*)\)", str(label)))

def match(label, vocab, allowed=None, exclude=None):
    """Best concept for a label. Returns (concept, score) or (None, 0)."""
    best_c, best_s = None, 0
    for v in variants(label):
        toks = set(v.split())
        for concept, syns in vocab.items():
            if allowed is not None and concept not in allowed:
                continue
            if exclude and toks & exclude.get(concept, set()):
                continue
            for syn in syns:
                st = syn.split()
                if len(st) == 1:
                    s = 1.5 if v == syn else 0
                elif set(st) <= toks:
                    s = len(st) + (0.5 if v == syn else 0)
                else:
                    s = 0
                if s > best_s:
                    best_c, best_s = concept, s
    return best_c, best_s

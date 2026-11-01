"""Countries and nationalities (name, ISO alpha-2, ISO alpha-3, nationality words, other-language names)."""
import re, unicodedata

def _n(s):
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c)).lower()
    return re.sub(r"[^a-z0-9]+", "", s)

# name, a2, a3, [nationality words], [names in other languages]
RAW = [
 ("Netherlands","NL","NLD",["Dutch","Netherlander"],["Niederlande","Paises Bajos","Holland","Nederland"]),
 ("Belgium","BE","BEL",["Belgian"],["Belgien","Belgica"]),
 ("Germany","DE","DEU",["German"],["Deutschland","Alemania"]),
 ("France","FR","FRA",["French"],["Frankreich","Francia"]),
 ("United Kingdom","GB","GBR",["British","English","Scottish","Welsh"],["Vereinigtes Konigreich","Reino Unido","Great Britain","UK"]),
 ("Ireland","IE","IRL",["Irish"],["Irland","Irlanda"]),
 ("Norway","NO","NOR",["Norwegian"],["Norwegen","Noruega"]),
 ("Sweden","SE","SWE",["Swedish"],["Schweden","Suecia"]),
 ("Denmark","DK","DNK",["Danish"],["Danemark","Dinamarca"]),
 ("Finland","FI","FIN",["Finnish"],["Finnland","Finlandia"]),
 ("Poland","PL","POL",["Polish"],["Polen","Polonia"]),
 ("Russia","RU","RUS",["Russian"],["Russland","Rusia","Russian Federation"]),
 ("Ukraine","UA","UKR",["Ukrainian"],["Ucrania"]),
 ("Philippines","PH","PHL",["Filipino","Philippine"],["Philippinen","Filipinas"]),
 ("India","IN","IND",["Indian"],["Indien"]),
 ("Indonesia","ID","IDN",["Indonesian"],["Indonesien"]),
 ("China","CN","CHN",["Chinese"],[]),
 ("Myanmar","MM","MMR",["Burmese"],["Burma","Birmania"]),
 ("Vietnam","VN","VNM",["Vietnamese"],["Viet Nam"]),
 ("Thailand","TH","THA",["Thai"],["Tailandia"]),
 ("Croatia","HR","HRV",["Croatian"],["Kroatien","Croacia"]),
 ("Romania","RO","ROU",["Romanian"],["Rumanien","Rumania"]),
 ("Bulgaria","BG","BGR",["Bulgarian"],["Bulgarien"]),
 ("Greece","GR","GRC",["Greek"],["Griechenland","Grecia"]),
 ("Turkey","TR","TUR",["Turkish"],["Turkei","Turquia"]),
 ("Spain","ES","ESP",["Spanish"],["Spanien","Espana"]),
 ("Portugal","PT","PRT",["Portuguese"],[]),
 ("Italy","IT","ITA",["Italian"],["Italien","Italia"]),
 ("Estonia","EE","EST",["Estonian"],["Estland"]),
 ("Latvia","LV","LVA",["Latvian"],["Lettland","Letonia"]),
 ("Lithuania","LT","LTU",["Lithuanian"],["Litauen","Lituania"]),
 ("Georgia","GE","GEO",["Georgian"],["Georgien"]),
 ("Sri Lanka","LK","LKA",["Sri Lankan"],[]),
 ("Pakistan","PK","PAK",["Pakistani"],[]),
 ("Bangladesh","BD","BGD",["Bangladeshi"],[]),
 ("Egypt","EG","EGY",["Egyptian"],["Agypten","Egipto"]),
 ("South Africa","ZA","ZAF",["South African"],["Sudafrika","Sudáfrica"]),
 ("Nigeria","NG","NGA",["Nigerian"],[]),
 ("Ghana","GH","GHA",["Ghanaian"],[]),
 ("Brazil","BR","BRA",["Brazilian"],["Brasilien","Brasil"]),
 ("Argentina","AR","ARG",["Argentine","Argentinian"],["Argentinien"]),
 ("Chile","CL","CHL",["Chilean"],[]),
 ("Peru","PE","PER",["Peruvian"],[]),
 ("Colombia","CO","COL",["Colombian"],["Kolumbien"]),
 ("Mexico","MX","MEX",["Mexican"],["Mexiko"]),
 ("United States","US","USA",["American"],["Vereinigte Staaten","Estados Unidos","United States of America"]),
 ("Canada","CA","CAN",["Canadian"],["Kanada"]),
 ("Australia","AU","AUS",["Australian"],["Australien"]),
 ("New Zealand","NZ","NZL",["New Zealander"],["Neuseeland","Nueva Zelanda"]),
 ("Japan","JP","JPN",["Japanese"],[]),
 ("South Korea","KR","KOR",["Korean","South Korean"],["Sudkorea","Corea del Sur","Korea Republic of"]),
 ("Singapore","SG","SGP",["Singaporean"],["Singapur"]),
 ("Malaysia","MY","MYS",["Malaysian"],[]),
 ("Montenegro","ME","MNE",["Montenegrin"],[]),
 ("Serbia","RS","SRB",["Serbian"],["Serbien"]),
 ("Slovenia","SI","SVN",["Slovenian"],["Slowenien","Eslovenia"]),
 ("Slovakia","SK","SVK",["Slovak"],["Slowakei","Eslovaquia"]),
 ("Czech Republic","CZ","CZE",["Czech"],["Tschechien","Chequia"]),
 ("Hungary","HU","HUN",["Hungarian"],["Ungarn","Hungria"]),
 ("Austria","AT","AUT",["Austrian"],["Osterreich"]),
 ("Switzerland","CH","CHE",["Swiss"],["Schweiz","Suiza"]),
 ("Iceland","IS","ISL",["Icelandic"],["Island","Islandia"]),
 ("Cyprus","CY","CYP",["Cypriot"],["Zypern","Chipre"]),
 ("Malta","MT","MLT",["Maltese"],[]),
 ("Morocco","MA","MAR",["Moroccan"],["Marokko","Marruecos"]),
 ("Kenya","KE","KEN",["Kenyan"],[]),
 ("Tanzania","TZ","TZA",["Tanzanian"],[]),
]
_index = {}
for name, a2, a3, dem, other in RAW:
    rec = {"name": name, "a2": a2, "a3": a3}
    for k in [name, a2, a3] + dem + other:
        _index.setdefault(_n(k), rec)

def lookup(text):
    """Return the country record for a country name, nationality word or ISO code, or None."""
    if text is None:
        return None
    return _index.get(_n(text))

"""Creates a filled-in test master (fake data) from the empty template."""
import sys, datetime, subprocess
from openpyxl import load_workbook
src, dst = sys.argv[1], sys.argv[2]
wb = load_workbook(src)
D = datetime.date
def kv(sheet, label, value):
    ws = wb[sheet]
    for r in range(4, ws.max_row + 1):
        if ws.cell(r, 2).value == label:
            ws.cell(r, 3).value = value; return
    raise KeyError(label)
for lab, v in {"Name of ship":"MV EXAMPLE STAR","IMO number":"9123456","Call sign":"PHAB","Flag State of ship":"Netherlands",
  "Name of master":"Capt. J. de Vries","Gross tonnage":12450,"Net tonnage":5200}.items(): kv("Ship", lab, v)
for lab, v in {"Voyage number":"2026-087","Arrival / Departure":"Arrival","Port of arrival/departure":"Antwerp",
  "Date and time of arrival/departure":datetime.datetime(2026,10,24,14,30),"Last port of call":"Rotterdam","Next port of call":"Hamburg"}.items(): kv("Voyage", lab, v)
for lab, v in {"Valid ISSC":"Yes","ISSC issuing company / authority":"Example Classification Society","ISSC issuer type":"RSO",
  "SSP on board":"Yes","ISSC expiry date":D(2029,5,15),"CSO given name":"Anna","CSO family name":"De Boer",
  "CSO phone 24/7":"+31 10 123 4567","CSO e-mail":"cso@exampleshipping.com","Company name":"Example Shipping B.V.",
  "Company IMO number":"1234567","Ship configuration":"Double hull","Condition of cargo tanks":"Empty","Ballast tanks":"Full"}.items(): kv("Security", lab, v)
crew = [
 ("JANSEN","Pieter Willem","Master","Dutch",D(1975,3,15),"Rotterdam","M","Passport","NLD123456","Netherlands",D(2029,6,20)),
 ("DE GROOT","Marco","Chief Officer","Dutch",D(1985,3,15),"Utrecht","M","Seaman's book","NLD223344","Netherlands",D(2028,1,10)),
 ("SANTOS","Maria Luz","Second Officer","Filipino",D(1990,7,2),"Manila","F","Passport","PH778899","Philippines",D(2030,2,2)),
 ("IVANOV","Sergey","Chief Engineer","Ukrainian",D(1972,11,30),"Odessa","M","Passport","UA556677","Ukraine",D(2027,9,9)),
 ("KOWALSKI","Jan","Second Engineer","Polish",D(1988,5,5),"Gdansk","M","Passport","PL112233","Poland",D(2031,4,4)),
 ("REYES","Carlos","Bosun","Filipino",D(1980,1,12),"Cebu","M","Seaman's book","PH445566","Philippines",D(2029,8,8)),
 ("CRUZ","Juan","Able Seaman","Filipino",D(1992,9,9),"Davao","M","Seaman's book","PH998877","Philippines",D(2028,3,3)),
 ("SANTOS","Pedro","AB","Filipino",D(1995,2,14),"Manila","M","Seaman's book","PH334455","Philippines",D(2030,6,6)),
 ("PETROV","Alexei","Oiler","Russian",D(1983,6,21),"Murmansk","M","Passport","RU101010","Russia",D(2026,12,31)),
 ("MULLER","Hans","Electrician","German",D(1979,4,4),"Hamburg","M","Passport","DE202020","Germany",D(2029,10,10)),
 ("LARSEN","Erik","Cook","Norwegian",D(1968,12,1),"Bergen","M","Passport","NO303030","Norway",D(2027,5,5)),
 ("NGUYEN","Linh","Steward","Vietnamese",D(1993,3,3),"Hanoi","F","Passport","VN404040","Vietnam",D(2028,11,11)),
 ("DE BOER","Kees","Deck Cadet","Dutch",D(2001,8,8),"Delft","M","Passport","NLD505050","Netherlands",D(2031,8,8)),
 ("ALVAREZ","Diego","Ordinary Seaman","Spanish",D(1997,10,20),"Vigo","M","Passport","ES606060","Spain",D(2030,1,1)),
 ("KHAN","Imran","Wiper","Pakistani",D(1989,2,28),"Karachi","M","Passport","PK707070","Pakistan",D(2029,12,12)),
 ("SILVA","Rui","Motorman","Portuguese",D(1986,7,17),"Lisbon","M","Passport","PT808080","Portugal",D(2028,7,7)),
]
ws = wb["Crew_List"]
for i, p in enumerate(crew):
    r = 6 + i
    for j, v in enumerate(p, start=2):
        ws.cell(r, j).value = v
    # columns: B family, C given, D rank, E nationality, F dob, G pob, H gender, I id type, J id no, K issuing state, L expiry
ws2 = wb["Ship_Stores"]
stores = [("Bonded - alcohol","Whisky 70cl",12,"liters"),("Bonded - alcohol","Vodka",8,"liters"),("Bonded - alcohol","Beer cans",120,"liters"),
          ("Bonded - alcohol","Red wine",30,"liters"),("Bonded - tobacco","Cigarettes",1800,"pieces"),("Bonded - tobacco","Tobacco",4,"kg"),
          ("Fuel","HFO",250000,"liters"),("Fuel","MGO",45000,"liters"),("Lubricating oil","Lub oil",6000,"liters"),
          ("Bonded - alcohol","Gin",24,"bottles")]
for i, (cat, name, q, u) in enumerate(stores):
    r = 6 + i
    ws2.cell(r,1).value, ws2.cell(r,2).value, ws2.cell(r,3).value, ws2.cell(r,4).value, ws2.cell(r,5).value = cat, name, q, u, "Store"
wp = wb["Port_History"]
ports = [("Gdynia","Poland","PLGDY",D(2026,10,1),D(2026,10,3)),("Bremerhaven","Germany","DEBRV",D(2026,10,5),D(2026,10,7)),
 ("Le Havre","France","FRLEH",D(2026,10,9),D(2026,10,11)),("Southampton","United Kingdom","GBSOU",D(2026,10,13),D(2026,10,14)),
 ("Felixstowe","United Kingdom","GBFXT",D(2026,10,16),D(2026,10,17)),("Hamburg","Germany","DEHAM",D(2026,10,19),D(2026,10,20)),
 ("Rotterdam","Netherlands","NLRTM",D(2026,10,22),D(2026,10,23))]
for i, (p, c, l, a, d) in enumerate(ports):
    r = 6 + i
    for j, v in enumerate((p, c, l, a, d, f"{l}-0001", 1, "None"), start=1):
        wp.cell(r, j).value = v
ww = wb["Waste"]
ww["C4"], ww["C5"], ww["C6"] = "Rotterdam", D(2026,10,22), "Hamburg"
def waste(code, vals):
    for r in range(9, ww.max_row):
        if str(ww.cell(r,1).value or "") == code and ww.cell(r,1).value:
            for j, v in enumerate(vals, start=3):
                if v is not None: ww.cell(r, j).value = v
            return
waste("101", (2.5, 10, 4.0, 3.0, 1.5)); waste("102", (4.0, 20, 9.5, 5.0, None)); waste("502", (0.3, 2, 0.5, 0.4, 0.2))
waste("501", (0.2, 1, 0.3, 0.1, None)); waste("401", (5.0, 15, 2.0, 6.0, None))
wb.save(dst)

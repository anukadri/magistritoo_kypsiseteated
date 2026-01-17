""" 29. september 2025
Anu Kadri Uustalu magistritöö programm
aitab vastata küsimustele
                                kas veebilehe haldaja mainib ennast küpsiseteates?
                                kas küpsiseid kasutab veebileht või inimene
Programm loeb sisse MA töö andmestiku,

väljund: .txt fail, tabuleeritud andmetega"""

from openpyxl.workbook import Workbook as wb
from openpyxl import load_workbook

# avab faili
path = r"Uustalu_MA_andmestik.xlsx" # andmestiku faili path
wbAndmestik = load_workbook(path)
wsAndmestik = wbAndmestik["Andmestik"] #worksheet

# millises veerus on milline info

tervikTekstiLeidja = "A"
veergID = "I"
veergSektor = "BK"
veergDistantseerimine = "CF" #kas kasutatakse "meie", "see veebileht", firma nimi vms
veergKesKasutab = "CG" #"veebileht kasutab küpsiseid", "meie kasutame küpsiseid", muu


# teeb tühjad sõnastikud, et sinna lugeda vastavate kriteeriumidega lahtri sisud ja väärtuseks nende kogused
"""kokkuDistantseerimine = {}
eraDistantseerimine = {}
avalikDistantseerimine = {}

kokkuKasutamine = {}
eraKasutamine = {}
avalikKasutamine = {}"""

distantseerimineDict = {}
kasutamineDict = {}


rida = 1
while True:
    rida += 1
    lahter = "".join((tervikTekstiLeidja, str(rida))) # leiab lahtri, kust teksti leida, nt A1
    tervikTekst = wsAndmestik[lahter].value
    
    if tervikTekst == None: # kood lõppeb ära, kui tuleb ette tühi lahter
        break
    if tervikTekst == "na": # programm jätab vahele kõik "na" read ehk lause mitte tervikteksti read
        continue
    
    
    idKood = wsAndmestik["".join((veergID, str(rida)))].value
    
    #print(idKood)
    
    sisuSektor = wsAndmestik["".join((veergSektor, str(rida)))].value
    if sisuSektor == "era":
        sektor = "era"
    elif sisuSektor == "avalik":
        sektor = "avalik"
    else:
        print(sisuSektor)
    
    distantseerimine = wsAndmestik["".join((veergDistantseerimine, str(rida)))].value
    kesKasutab = wsAndmestik["".join((veergKesKasutab, str(rida)))].value
    
    if distantseerimine not in distantseerimineDict.keys():
        distantseerimineDict[distantseerimine] = {"Kokku": 1}
        if sektor == "era":
            distantseerimineDict[distantseerimine]["Era"] = 1
        elif sektor == "avalik":
            distantseerimineDict[distantseerimine]["Avalik"] = 1

    else:
        distantseerimineDict[distantseerimine]["Kokku"] += 1
        try:
            if sektor == "era":
                distantseerimineDict[distantseerimine]["Era"] += 1
        except:
            distantseerimineDict[distantseerimine]["Era"] = 1
        try:
            if sektor == "avalik":
                distantseerimineDict[distantseerimine]["Avalik"] += 1
        except:
            distantseerimineDict[distantseerimine]["Avalik"] = 1
            
            
    if kesKasutab not in kasutamineDict.keys():
        kasutamineDict[kesKasutab] = {"Kokku": 1}
        if sektor == "era":
            kasutamineDict[kesKasutab]["Era"] = 1
        elif sektor == "avalik":
            kasutamineDict[kesKasutab]["Avalik"] = 1

    else:
        kasutamineDict[kesKasutab]["Kokku"] += 1
        try:
            if sektor == "era":
                kasutamineDict[kesKasutab]["Era"] += 1
        except:
            kasutamineDict[kesKasutab]["Era"] = 1
        try:
            if sektor == "avalik":
                kasutamineDict[kesKasutab]["Avalik"] += 1
        except:
            kasutamineDict[kesKasutab]["Avalik"] = 1

  



fail = open("Kysimus5_distantseerimine_kasutamine.txt", "w", encoding="UTF-8")

fail.write("Veebilehtede nimetamine küpsiseteadetes\n")
fail.write("\tEra\tAvalik\tKokku\n")

for key, value in distantseerimineDict.items():
    try:
        era = value["Era"]
    except:
        era = 0
    try:
        avalik = value["Avalik"]
    except:
        avalik = 0
    try:
        kokku = value["Kokku"]
    except:
        kokku = 0
    fail.write(f"{key}\t{era}\t{avalik}\t{kokku}\n")
    
fail.write("\n\nKes kasutab küpsiseid, veebileht või selle haldaja?\n")
fail.write("\tEra\tAvalik\tKokku\n")
for key, value in kasutamineDict.items():
    try:
        era = value["Era"]
    except:
        era = 0
    try:
        avalik = value["Avalik"]
    except:
        avalik = 0
    try:
        kokku = value["Kokku"]
    except:
        kokku = 0
    fail.write(f"{key}\t{era}\t{avalik}\t{kokku}\n")


fail.close()
print("Valmis!")

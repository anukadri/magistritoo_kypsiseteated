"""
25. september 2025
Anu Kadri Uustalu magistritöö programm
aitab vastata küsimustele kas sisu- ja nuputekstid sobivad omavahel grammatiliselt ja pragmaatiliselt 


väljund: .txt fail
"""

from openpyxl.workbook import Workbook as wb
from openpyxl import load_workbook

# avab faili
path = r"Uustalu_MA_andmestik.xlsx" # andmestiku faili path
wbAndmestik = load_workbook(path)
wsAndmestik = wbAndmestik["Andmestik"] #worksheet

# millises veerus on milline info

tervikTekstiLeidja = "A"
veergID = "I"
veergIdNimi = "CR"
veergSektor = "BK"
veergVerb = "CN"

veergVerb1sg = "CI"
veergImperatiiv = "CJ"
veergEitus = "CK"
veergNimis6na = "CL"
veergMuu = "CM"

arvutiS6nad = ["nõustuma", "lubama", "keelduma", "kohandama",
                       "muutma", "lugema", "kinnitama",
                       "vaatama", "haldama", "nõus olema", "sulgema",
                       "salvestama", "tagasi lükkama", "valima", "keelama",
                       "seadistama", "loobuma", "personaliseerima",
                       "määrama", "aru saama", "tutvuma", "soovima",
                       "pressima", "manageerima", "aktsepteerima", "peitma",
                       "kohendama", "mitu verbi"] # käskivas kõneviisis, puuduvad näitama ja kuvama

kasutajaS6nadEndale = ["nõustuma", "lubama", "keelduma", "kohandama",
                       "muutma", "lugema", "kinnitama",
                       "vaatama", "haldama", "nõus olema",
                       "salvestama", "tagasi lükkama", "valima", "keelama",
                       "seadistama", "loobuma", "personaliseerima",
                       "määrama", "aru saama", "tutvuma", "soovima",
                       "pressima", "manageerima", "aktsepteerima",
                       "kohendama", "mitu verbi"] # SG1 või eitus, puuduvad näitama ja kuvama, sulgema, peitma

kasutajaS6nadArvutile = ["näitama", "kuvama", "sulgema", "peitma"] # käskiv kõneviis

tekstideS6nastik = {}
arvutiSuhtlebKasutajaga = []
kasutajaSuhtlebArvutiga = []
ebareeglip2rased = []
ylej22nud = []
ylej22nudS6nastik = {}


j2tabVahele2 = ["eant_ktea1",
               "emm_ktea1", "eramets_ktea1", "etag_ktea1",
               "imag_ktea2", "jan_ktea2", "pk_ktea1",
               "pmv_ktea1", "polvhai_ktea1", "proto_ktea3",
               "res_ktea1", "say_ktea1", "woole_ktea1", "petcity_ktea1"]

j2tabVahele = ["asa_ktea1", "tktk_ktea1", "eant_ktea1",
               "emm_ktea1", "eramets_ktea1", "etag_ktea1",
               "imag_ktea2", "jan_ktea2", "pk_ktea1",
               "pmv_ktea1", "polvhai_ktea1", "proto_ktea3",
               "res_ktea1", "say_ktea1", "woole_ktea1", "petcity_ktea1"] # jäävad välja, sest nupud on ainult NP või MUU

j2tabVahele.extend(["ad_ktea1", "eam_ktea1", "etdm_ktea1",
                   "lauam_ktea1", "prisma_ktea6", "rji_ktea1"]) #jäävad välja, sest nendes on nupud sulgema ja peitma, mille märgendan käsitsi

nuppePole = ["ahhaa_booking_ktea1", "epp_ktea1", "fertilitas_ktea_2",
                    "now_tervik_ktea1", "seto_tervik_ktea1"] # jäävad välja, sest nuppe ei ole

tekstVale = ["asa_ktea1", "tktk_ktea1"] # jäävad vahele, sest nende tekst on kahest erinevast vaatepunktist




rida = 1
while True:
    rida += 1
    verb = wsAndmestik["".join((veergVerb, str(rida)))].value # leiab lahtri, kust teksti leida, nt A1
    
    if verb == None: # kood lõppeb ära, kui tuleb ette tühi lahter
        break
    if verb == "na": # programm jätab vahele kõik "na" read ehk lause, mitte tervikteksti read
        continue
    
    
    idKood = wsAndmestik["".join((veergID, str(rida)))].value
    
    #print(idKood)
    
    idNimi = str(wsAndmestik["".join((veergIdNimi, str(rida)))].value).strip()
    sektor = wsAndmestik["".join((veergSektor, str(rida)))].value
    
    verb1sg = wsAndmestik["".join((veergVerb1sg, str(rida)))].value
    imperatiiv = wsAndmestik["".join((veergImperatiiv, str(rida)))].value
    eitus = wsAndmestik["".join((veergEitus, str(rida)))].value
    nimis6na = wsAndmestik["".join((veergNimis6na, str(rida)))].value
    muu = wsAndmestik["".join((veergMuu, str(rida)))].value
    
    if idNimi in j2tabVahele: # nende sisutekstid on kahest vaatepunktist või ainult np ja muu tekstidega seega siia arvestusse ei tule
        continue

    
    # see sõna on arvuti poolt kasutajale suunatud
    if verb in arvutiS6nad and imperatiiv == "jah":
        if not idNimi:
            continue

        if idNimi not in tekstideS6nastik:
            tekstideS6nastik[idNimi] = []

        if "Arvuti -> kasutaja" not in tekstideS6nastik[idNimi]:
            tekstideS6nastik[idNimi].append("Arvuti -> kasutaja")
            
        
    # see sõna on kasutaja poolt arvutile suunatud ning kasutaja ütleb seda enda kohta
    elif verb in kasutajaS6nadEndale and verb1sg == "PR":
        if not idNimi:
            continue

        if idNimi not in tekstideS6nastik:
            tekstideS6nastik[idNimi] = []

        if "Kasutaja (enda kohta) -> arvuti" not in tekstideS6nastik[idNimi]:
            tekstideS6nastik[idNimi].append("Kasutaja (enda kohta) -> arvuti")

            
            
    elif verb == "aru saama" and muu == "sain aru":
        try:
            if "Kasutaja (enda kohta) -> arvuti" not in tekstideS6nastik[idNimi]:
                if idNimi in tekstideS6nastik.keys():
                    tekstideS6nastik[idNimi].append("Kasutaja (enda kohta) -> arvuti")
        except:
            tekstideS6nastik[idNimi] = ["Kasutaja (enda kohta) -> arvuti"]
            
    elif verb in kasutajaS6nadEndale and eitus == "jah":
        try:
            if "Kasutaja (enda kohta) -> arvuti" not in tekstideS6nastik[idNimi]:
                if idNimi in tekstideS6nastik.keys():
                    tekstideS6nastik[idNimi].append("Kasutaja (enda kohta) -> arvuti")
        except:
            tekstideS6nastik[idNimi] = ["Kasutaja (enda kohta) -> arvuti"]
        
    # see sõna on kasutaja poolt arvutile suunatud ning kasutaja ütleb seda arvutile
    elif verb in kasutajaS6nadArvutile and imperatiiv == "jah":
        try:
            if "Kasutaja -> arvuti (arvutile)" not in tekstideS6nastik[idNimi]:
                if idNimi in tekstideS6nastik.keys():
                    tekstideS6nastik[idNimi].append("Kasutaja -> arvuti (arvutile)")
        except:
            tekstideS6nastik[idNimi] = ["Kasutaja -> arvuti (arvutile)"]
    # see sõna ei vasta seatud reeglitega
    elif verb != "NP" and verb != "MUU":
        if idNimi not in ebareeglip2rased:
            ebareeglip2rased.append(idNimi)
        try:
            if "Ebareeglip2rane" not in tekstideS6nastik[idNimi]:
                if idNimi in tekstideS6nastik.keys():
                    tekstideS6nastik[idNimi].append("Ebareeglip2rane")
        except:
            tekstideS6nastik[idNimi] = ["Ebareeglip2rane"]

            

#ylej22nudV22rtusteKaupa = {}
for v6ti, v22rtus in tekstideS6nastik.items():
    if v22rtus == ["Arvuti -> kasutaja"]:
        arvutiSuhtlebKasutajaga.append(v6ti)
    elif v22rtus == ["Kasutaja -> arvuti (arvutile)"]:
        kasutajaSuhtlebArvutiga.append(v6ti)
    elif v22rtus == ["Kasutaja (enda kohta) -> arvuti", "Kasutaja -> arvuti (arvutile)"] or v22rtus == ["Kasutaja -> arvuti (arvutile)", "Kasutaja (enda kohta) -> arvuti"]:
        kasutajaSuhtlebArvutiga.append(v6ti)
    elif v22rtus == ["Kasutaja (enda kohta) -> arvuti"]:
        kasutajaSuhtlebArvutiga.append(v6ti)
    elif "Ebareeglip2rane" in v22rtus:
        if v6ti not in ebareeglip2rased:
            ebareeglip2rased.append(v6ti)
        #print(v6ti)
    else:
        ylej22nud.append(v6ti)
        ylej22nudS6nastik[v6ti] = v22rtus
        #print(v22rtus)
        
    
    """
    if not tuple(v22rtus) in ylej22nudS6nastik:
        ylej22nudS6nastik[tuple(v22rtus)] = [v6ti]  
    else:
        ylej22nudS6nastik[tuple(v22rtus)].append(v6ti)"""
        
### lisan käsitsi märgendatud asjad
v6ibM6lematPidi = ["eam_ktea1", "etdm_ktea1", "rji_ktea1"]
kasutajaSuhtlebArvutiga.extend(["ad_ktea1", "lauam_ktea1", "prisma_ktea6"])



#print(tekstideS6nastik)

print("Arvuti suhtleb kasutajaga", len(arvutiSuhtlebKasutajaga))
print("Kasutaja suhtleb arvutiga", len(kasutajaSuhtlebArvutiga))
print("Võib mõlemat pidi", len(v6ibM6lematPidi))
print("Ülejäänud", len(ylej22nud))
print("Ebareeglipärased verb + pööre kombod", len(ebareeglip2rased))
#print(ebareeglip2rased)




fail = open("Kysimus6_sisujanupp.txt", "w", encoding="UTF-8")


fail.write("Kas arvuti suhtleb kasutajaga või kasutaja arvutiga ehk küsimus nr 6\n\n")
fail.write(f"Selliseid teateid, kus arvuti suhtleb selgelt kasutajaga on {len(arvutiSuhtlebKasutajaga)}\n")
fail.write(f"Selliseid teateid, kus kasutaja suhtleb selgelt arvutiga on {len(kasutajaSuhtlebArvutiga)}\n")
fail.write(f"Ebamääraseid teateid, kus kaks varianti on läbisegi on {len(ylej22nud)}\n\n")

fail.write("Need on järgnevad:\n")
for elem in ylej22nud:
    fail.write(f"{elem}\n")
    
fail.write("\n\nLisaks esineb ka 6 ebareeglipärast nuputeksti. Need leiduvad järgnevates teadetes:\n")
for elem in ebareeglip2rased:
    fail.write(f"{elem}\n")

fail.close()
print("Valmis!")


fail = open("Kysimus6_sisujanupp_tabel.txt", "w", encoding="UTF-8")


fail.write("Id kood\tkategooria\n")
for elem in arvutiSuhtlebKasutajaga:
    fail.write(f"{elem}\tarvuti suhtleb kasutajaga\n")
for elem in kasutajaSuhtlebArvutiga:
    fail.write(f"{elem}\tkasutaja suhtleb arvutiga\n")
for elem in ylej22nud:
    fail.write(f"{elem}\tkaks varianti läbisegi\n")
for elem in v6ibM6lematPidi:
    fail.write(f"{elem}\tSaab mõlemat pidi tõlgendada\n")
for elem in ebareeglip2rased:
    fail.write(f"{elem}\tEbareeglipärased\n")
for elem in j2tabVahele2:
    fail.write(f"{elem}\tJääb vahele: NP või MUU\n")
for elem in nuppePole:
    fail.write(f"{elem}\tJääb vahele: nuppe pole\n")
for elem in tekstVale:
    fail.write(f"{elem}\tJääb vahele: sisutekst on kahest küljest\n")



fail.close()
print("Valmis!")

"""
komplekt1IMP = []
komplekt2 = []
komplekt3k6ik = []
for v6ti, v22rtus in ylej22nudV22rtusteKaupa.items():
    if v6ti == ('Kasutaja -> arvuti (arvutile)', 'Arvuti -> kasutaja') or v6ti == ('Arvuti -> kasutaja', 'Kasutaja -> arvuti (arvutile)'):
        komplekt1IMP.append(v6ti)
    if v6ti == ('Arvuti -> kasutaja', 'Kasutaja (enda kohta) -> arvuti') or v6ti == ('Kasutaja (enda kohta) -> arvuti', 'Arvuti -> kasutaja'):
        komplekt2.append(v6ti)
    if v6ti ==('Arvuti -> kasutaja', 'Kasutaja -> arvuti (arvutile)', 'Kasutaja (enda kohta) -> arvuti') or v6ti == ('Kasutaja -> arvuti (arvutile)', 'Arvuti -> kasutaja', 'Kasutaja (enda kohta) -> arvuti'):
        komplekt3k6ik.append(v6ti)
    

print(len(komplekt1IMP))
print(len(komplekt2))
print(len(komplekt3k6ik))"""
"""
for v6ti in ylej22nudV22rtusteKaupa.keys():
    print(v6ti)
    
for v6ti, v22rtus in ylej22nudV22rtusteKaupa.items():
    if v6ti == ('Arvuti -> kasutaja',):
        print(v22rtus)"""
        



verbidVormides = []

fail = open("Kysimus6_ylej22nud_s6navormid.txt", "w", encoding="UTF-8")
fail.write("Id nimi\tVerb\tVorm\n")

rida = 1
while True:
    rida += 1
    idNimi = str(wsAndmestik["".join((veergIdNimi, str(rida)))].value).strip()
    verb = wsAndmestik["".join((veergVerb, str(rida)))].value # leiab lahtri, kust teksti leida, nt A1
    verb1sg = wsAndmestik["".join((veergVerb1sg, str(rida)))].value
    imperatiiv = wsAndmestik["".join((veergImperatiiv, str(rida)))].value
    eitus = wsAndmestik["".join((veergEitus, str(rida)))].value
    muu = wsAndmestik["".join((veergMuu, str(rida)))].value
    
    if verb == None: # kood lõppeb ära, kui tuleb ette tühi lahter
        break
    if verb == "NP":
        continue
    if verb == "MUU":
        continue
    if verb == "na": # programm jätab vahele kõik "na" read ehk lause, mitte tervikteksti read
        continue
    
    idNimi = str(wsAndmestik["".join((veergIdNimi, str(rida)))].value).strip()

    if idNimi not in ylej22nud:
        continue
    else:
        if verb1sg == None:
            verb1sg = ""
        if imperatiiv == None:
            imperatiiv = ""
        if eitus == None:
            eitus = ""
        if verb1sg == "PR":
            verb1sg = "_1sgPR"
        if imperatiiv == "jah":
            imperatiiv = "_IMP2sg"
        if eitus == "jah":
            eitus = "_NEG"
        if muu == None:
            muu = ""
        elif muu != None:
            muu = "MUU"
        #verbidVormides.append(verb+verb1sg+imperatiiv+eitus+muu)
        
        #for elem in verbidVormides:
        fail.write(f"{idNimi}\t{verb}\t{verb1sg}{imperatiiv}{eitus}{muu}\n")

fail.close()
print("Valmis!")



fail = open("Kysimus6_tekstideJaotus.txt", "w", encoding="UTF-8")
fail.write("ARVUTI SUHTLEB KASUTAJAGA\n")
fail.write(f"{len(arvutiSuhtlebKasutajaga)}\n")
for elem in arvutiSuhtlebKasutajaga:
    fail.write(elem)
    fail.write("\n")
    
fail.write("\n\n\n")
    
fail.write("KASUTAJA SUHTLEB ARVUTIGA\n")
fail.write(f"{len(kasutajaSuhtlebArvutiga)}\n")
for elem in kasutajaSuhtlebArvutiga:
    fail.write(elem)
    fail.write("\n")
    
fail.write("\n\n\n")
    
fail.write("EBAREEGLIPÄRASED\n")
fail.write(f"{len(ebareeglip2rased)}\n")
for elem in ebareeglip2rased:
    fail.write(elem)
    fail.write("\n")
    
fail.write("\n\n\n")
    
fail.write("ÜLEJÄÄNUD\n")
fail.write(f"{len(ylej22nud)}\n")
for elem in ylej22nud:
    fail.write(elem)
    fail.write("\n")

fail.close()




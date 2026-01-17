""" 26. september 2025
Anu Kadri Uustalu magistritöö programm
aitab vastata küsimustele milliseid viisakuse võtteid kasutatakse mis sektorites ja kui palju

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
veergTeieSina = "BM"
veergSuurAlgust2ht = "BN"
veergTingivK6neviis = "BR"
veergHyyum2rgid = "BQ"
veergEmotikon = "AZ"
veergPalun = "P"

#veergTekstityyp = "J" #peab olema kas sisutekst või pealkiri


# maatriksi ridades esimene on distantseeriv, teine lähendav
teieSinaV22rtused = [[0, 0], #teie, sina kokku
                     [0, 0], # erasektor
                     [0, 0]] #avalik sektor
suurT2htV22rtused = [[0, 0], #suur täht, väike täht
                     [0, 0], 
                     [0, 0]] 
tingivK6neviisV22rtused = [[0, 0], #ei ole tingiv, on tingiv
                     [0, 0], 
                     [0, 0]] 
hyym2rgidV22rtused = [[0, 0], #ei kasutata hüüumärke, kasutatakse hüüumärke
                     [0, 0], 
                     [0, 0]]
emotikonidV22rtused = [[0, 0], #ei kasutata emotikone/emojisid, kasutatakse emotikone/emojisid
                     [0, 0], 
                     [0, 0]]
palunV22rtused = [[0, 0], #palun kasutatakse, palun ei kasutata
                     [0, 0], 
                     [0, 0]]



rida = 1
while True:
    rida += 1
    lahter = "".join((tervikTekstiLeidja, str(rida))) # leiab lahtri, kust teksti leida, nt A1
    tervikTekst = wsAndmestik[lahter].value
    
    if tervikTekst == None: # kood lõppeb ära, kui tuleb ette tühi lahter
        break
    if tervikTekst == "na": # programm jätab vahele kõik "na" read ehk lause, mitte tervikteksti read
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
    
    teieSina = wsAndmestik["".join((veergTeieSina, str(rida)))].value
    suurAlgust2ht = wsAndmestik["".join((veergSuurAlgust2ht, str(rida)))].value
    tingivK6neviis = wsAndmestik["".join((veergTingivK6neviis, str(rida)))].value
    hyyum2rgid = wsAndmestik["".join((veergHyyum2rgid, str(rida)))].value
    emotikon = wsAndmestik["".join((veergEmotikon, str(rida)))].value
    palun = wsAndmestik["".join((veergPalun, str(rida)))].value
    
    
    if teieSina == "teie":
        teieSinaV22rtused[0][0] += 1
        if sektor == "era":
            teieSinaV22rtused[1][0] += 1
        if sektor == "avalik":
            teieSinaV22rtused[2][0] += 1
    if teieSina == "sina":
        teieSinaV22rtused[0][1] += 1
        if sektor == "era":
            teieSinaV22rtused[1][1] += 1
        if sektor == "avalik":
            teieSinaV22rtused[2][1] += 1
            
    if suurAlgust2ht == "suur":
        suurT2htV22rtused[0][0] += 1
        if sektor == "era":
            suurT2htV22rtused[1][0] += 1
        if sektor == "avalik":
            suurT2htV22rtused[2][0] += 1
    if suurAlgust2ht == "väike":
        suurT2htV22rtused[0][1] += 1
        if sektor == "era":
            suurT2htV22rtused[1][1] += 1
        if sektor == "avalik":
            suurT2htV22rtused[2][1] += 1
        
    if tingivK6neviis == "na":
        tingivK6neviisV22rtused[0][0] += 1
        if sektor == "era":
            tingivK6neviisV22rtused[1][0] += 1
        if sektor == "avalik":
            tingivK6neviisV22rtused[2][0] += 1
    if tingivK6neviis == "tingiv":
        tingivK6neviisV22rtused[0][1] += 1
        if sektor == "era":
            tingivK6neviisV22rtused[1][1] += 1
        if sektor == "avalik":
            tingivK6neviisV22rtused[2][1] += 1
        
    if hyyum2rgid == "na":
        hyym2rgidV22rtused[0][0] += 1
        if sektor == "era":
            hyym2rgidV22rtused[1][0] += 1
        if sektor == "avalik":
            hyym2rgidV22rtused[2][0] += 1
    if hyyum2rgid != "na":
        hyym2rgidV22rtused[0][1] += 1
        if sektor == "era":
            hyym2rgidV22rtused[1][1] += 1
        if sektor == "avalik":
            hyym2rgidV22rtused[2][1] += 1
            
    if emotikon == "na":
        emotikonidV22rtused[0][0] += 1
        if sektor == "era":
            emotikonidV22rtused[1][0] += 1
        if sektor == "avalik":
            emotikonidV22rtused[2][0] += 1
    if emotikon == "jah":
        emotikonidV22rtused[0][1] += 1
        if sektor == "era":
            emotikonidV22rtused[1][1] += 1
        if sektor == "avalik":
            emotikonidV22rtused[2][1] += 1
            
    if palun == "jah":
        palunV22rtused[0][0] += 1
        if sektor == "era":
            palunV22rtused[1][0] += 1
        if sektor == "avalik":
            palunV22rtused[2][0] += 1
    if palun == "ei":
        palunV22rtused[0][1] += 1
        if sektor == "era":
            palunV22rtused[1][1] += 1
        if sektor == "avalik":
            palunV22rtused[2][1] += 1
        
print(teieSinaV22rtused)
print(suurT2htV22rtused)
print(tingivK6neviisV22rtused)
print(hyym2rgidV22rtused)
print(emotikonidV22rtused)
print(palunV22rtused)


fail = open("Kysimus3_4_viisakusvormid.txt", "w", encoding="UTF-8")

fail.write("Sektorite viisakusvormide kogused Exceli stacked barcharti jaoks\n")
fail.write(" \t \tdistantseeriv\tlähendav\n")
fail.write("Erasektor\n")
fail.write(f"\tPronoomen (teie-sina)\t{teieSinaV22rtused[1][0]}\t{teieSinaV22rtused[1][1]}\n")
fail.write(f"\tPronoomeni algustäht (suur-väike)\t{suurT2htV22rtused[1][0]}\t{suurT2htV22rtused[1][1]}\n")
fail.write(f"\tKõneviis (muu-tingiv)\t{tingivK6neviisV22rtused[1][0]}\t{tingivK6neviisV22rtused[1][1]}\n")
fail.write(f"\tHüümärgid (ei ole-on)\t{hyym2rgidV22rtused[1][0]}\t{hyym2rgidV22rtused[1][1]}\n")
fail.write(f"\tEmotikonid/emojid (ei ole-on)\t{emotikonidV22rtused[1][0]}\t{emotikonidV22rtused[1][1]}\n")
fail.write(f"\tPalun (on-ei ole)\t{palunV22rtused[1][0]}\t{palunV22rtused[1][1]}\n")


fail.write("\n")
fail.write("Avalik sektor\n")
fail.write(f" \tPronoomen (teie-sina)\t{teieSinaV22rtused[2][0]}\t{teieSinaV22rtused[2][1]}\n")
fail.write(f" \tPronoomeni algustäht (suur-väike)\t{suurT2htV22rtused[2][0]}\t{suurT2htV22rtused[2][1]}\n")
fail.write(f" \tKõneviis (muu-tingiv)\t{tingivK6neviisV22rtused[2][0]}\t{tingivK6neviisV22rtused[2][1]}\n")
fail.write(f" \tHüümärgid (ei ole-on)\t{hyym2rgidV22rtused[2][0]}\t{hyym2rgidV22rtused[2][1]}\n")
fail.write(f" \tEmotikonid/emojid (ei ole-on)\t{emotikonidV22rtused[2][0]}\t{emotikonidV22rtused[2][1]}\n")
fail.write(f" \tPalun (on-ei ole)\t{palunV22rtused[2][0]}\t{palunV22rtused[2][1]}\n")

fail.close()
print("Valmis!")





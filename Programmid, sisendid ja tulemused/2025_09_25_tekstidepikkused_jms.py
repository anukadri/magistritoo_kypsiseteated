""" 25. september 2025
Anu Kadri Uustalu magistritöö programm
aitab vastata küsimustele mis on küpsiseteadete keskmine pikkus, mediaanpikkus, mediaan nupp-linkide arv ja keskmine nupp-linkide arv kokku ning kummaski sektoris

väljund: .txt fail, tabuleeritud andmetega, mis on vajaliku struktuuriga, et sellest saaks sankey diagrammi teha"""

from openpyxl.workbook import Workbook as wb
from openpyxl import load_workbook
from statistics import mean
from statistics import median
from statistics import stdev

# avab faili
path = r"Uustalu_MA_andmestik.xlsx" # andmestiku faili path
wbAndmestik = load_workbook(path)
wsAndmestik = wbAndmestik["Andmestik"] #worksheet

# millises veerus on milline info

tervikTekstiLeidja = "A"
veergID = "I"
veergSektor = "BK"
veergNupplinkideArv = "BY"
veergTekstipikkus = "BS"


kokkuTekstipikkused = []
avalikTekstipikkused = []
eraTekstipikkused = []
kokkuNupplingid = []
avalikNupplingid = []
eraNupplingid = []

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
    
    nupplinkideArv = wsAndmestik["".join((veergNupplinkideArv , str(rida)))].value
    tekstipikkus= wsAndmestik["".join((veergTekstipikkus, str(rida)))].value
    
    kokkuNupplingid.append(nupplinkideArv)
    kokkuTekstipikkused.append(tekstipikkus)
    
    
    sisuSektor = wsAndmestik["".join((veergSektor, str(rida)))].value
    if sisuSektor == "era":
        eraNupplingid.append(nupplinkideArv)
        eraTekstipikkused.append(tekstipikkus)
    elif sisuSektor == "avalik":
        avalikNupplingid.append(nupplinkideArv)
        avalikTekstipikkused.append(tekstipikkus)
    else:
        print(sisuSektor)
        
        

kokkuTekstiKeskmine = round(mean(kokkuTekstipikkused), 2)
kokkuTekstiMediaan = median(kokkuTekstipikkused)
kokkuNupplinkKeskmine = round(mean(kokkuNupplingid), 2)
kokkuNupplinkMediaan = median(kokkuNupplingid)

eraTekstiKeskmine = round(mean(eraTekstipikkused), 2)
eraTekstiMediaan = median(eraTekstipikkused)
eraNupplinkKeskmine = round(mean(eraNupplingid), 2)
eraNupplinkMediaan = median(eraNupplingid)

avalikTekstiKeskmine = round(mean(avalikTekstipikkused), 2)
avalikTekstiMediaan = median(avalikTekstipikkused)
avalikNupplinkKeskmine = round(mean(avalikNupplingid), 2)
avalikNupplinkMediaan = median(avalikNupplingid)

kokkuSD = round(stdev(kokkuTekstipikkused))
eraSD = round(stdev(eraTekstipikkused))
avalikSD = round(stdev(avalikTekstipikkused))


fail = open("Kysimus1_pikkused.txt", "w", encoding="UTF-8")

fail.write(f"Kahe sektori tekstide keskmine pikkus on {kokkuTekstiKeskmine} ja mediaan on {kokkuTekstiMediaan}\n")
fail.write(f"Kahe sektori nupp-linkide keskmine arv on {kokkuNupplinkKeskmine} ja mediaan on {kokkuNupplinkMediaan}\n\n")

fail.write(f"Erasektori tekstide keskmine pikkus on {eraTekstiKeskmine} ja mediaan on {eraTekstiMediaan}\n")
fail.write(f"Erasektori nupp-linkide keskmine arv on {eraNupplinkKeskmine} ja mediaan on {eraNupplinkMediaan}\n\n")

fail.write(f"Avaliku sektori tekstide keskmine pikkus on {avalikTekstiKeskmine} ja mediaan on {avalikTekstiMediaan}\n")
fail.write(f"Avaliku sektori nupp-linkide keskmine arv on {avalikNupplinkKeskmine} ja mediaan on {avalikNupplinkMediaan}\n\n")

fail.write(f"Lühim küpsiseteade andmestikus on {min(kokkuTekstipikkused)} ja pikim {max(kokkuTekstipikkused)} tähemärki pikk.\n\n")

fail.write(f"Sektorite peale kokku on standardhälve {kokkuSD}, erasektoris {eraSD} ja avalikus sektoris {avalikSD}.")

fail.close()
print("Valmis!")





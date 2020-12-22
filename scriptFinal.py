import xlrd
import csv
import glob
import sys
import os
import pandas as pd 
import numpy as np 

fileEntree = []
fileSortie = []
txtfiles = []
for file in glob.glob("*.xls"):
    txtfiles.append(file)
for filefind in txtfiles:
    fileSplited = filefind.split('_')
    if len(fileSplited) > 1:
        if fileSplited[0] == 'RegistreEntreesDND':
            fileEntree.append(filefind)
        if fileSplited[0] == 'RegistreSortiesDND':
            fileSortie.append(filefind)
if len(fileEntree) !=2:
    sys.exit()
if len(fileSortie) !=2:
    sys.exit()

def csv_from_excel():
    wb = xlrd.open_workbook(fileEntree[0])
    print(wb.sheets())
    sh = wb.sheet_by_name('Sheet0')
    your_csv_file = open('entree_bois_vert.csv', 'w')
    wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

    for rownum in range(sh.nrows):
        wr.writerow(sh.row_values(rownum))

    your_csv_file.close()

def csv_from_excel2():
    wb = xlrd.open_workbook(fileEntree[1])
    print(wb.sheets())
    sh = wb.sheet_by_name('Sheet0')
    your_csv_file = open('entree_pradie.csv', 'w')
    wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

    for rownum in range(sh.nrows):
        wr.writerow(sh.row_values(rownum))

    your_csv_file.close()

# runs the csv_from_excel function:
csv_from_excel()
csv_from_excel2()


filename = 'entree_bois_vert.csv'
filename2 = 'entree_pradie.csv'

with open(filename2, 'r') as csvfile2:
    readerPradie = csv.reader(csvfile2)
    with open(filename, 'r') as csvfile:
        reader = csv.reader(csvfile)
        next(readerPradie)
        with open('dechetTriesEntree.csv','w+') as new_file:
            writer = csv.writer(new_file)
            FER = ["AGS BLANC ET PEINT","ALU COMPLEXE","ALU MELE","FERRAILLE","FERRAILLE SANS VALEUR","LAITON","PLATIN","INOX","CUIVRE","METAUX FERREUX","CABLES","ZINC","BATTERIES","PLAQUES ALUMINIUM"]
            DIB = ["DECHETS","DECHETS VIDES","BALAYAGE","GRAVATS DECHETS","GRAVATS DECHETS VIDES","MELANGE CARTON PAPIER","POLYSTYRENE","SELECTIF"]
            DECHETS_ULTIMES = ["DECHETS ULTIMES"]
            CARTON = ["CARTON","CARTON SANS VALEUR","GROS DE MAGASIN","JOURNAUX","ROGNURES DE CAISSERIE"]
            DEEE = ["DEEE","D3E VIDES","DEEE VIDES","CUMULUS","NEONS"]
            PAPIER = ["AFNOR","PAPIER","BROCHURE","CHEQUES BROYES","ARCHIVES","ARCHIVES VIDES","EXTRA CLAIR"]
            PNEUS = ["PNEU","PNEU PL VIDES","PNEUS SOUILLES"]
            GRAVAT = ["GRAVAT","GRAVATS VIDES"]
            DIS = ["DIS","EMBALLAGES SOUILLES PLASTIQUES","BIDONS PLASTIQUES VIDES SOUILLES","MASTIC COLLE PEINTURE"]
            BOIS = ["BOIS","BOIS VIDES","CAGETTES","PALETTES","BOIS PRE BROYE","PALETTES CASSEES","PALETTES VIDEES","SCIURE","SOUCHES VIDEES","SOUCHES"]
            VEGETAUX = ["VEGETAUX","VEGETAUX VIDES","PELOUSE VIDEE","VEGETAUX VIDES CEE"]
            PLASTIQUE = ["PLASTIQUE","PARE CHOCS","BIG BAG","BIG BIG A TRAITER"]
            for row in reader:
                if row:
                    if row[3] in FER:
                        row[3] = "FER"
                    if row[3] in DIB:
                        row[3] = "DIB"
                    if row[3] in DECHETS_ULTIMES:
                        row[3] = "DECHETS_ULTIMES"
                    if row[3] in CARTON:
                        row[3] = "CARTON"
                    if row[3] in DEEE:
                        row[3] = "DEEE"
                    if row[3] in PAPIER:
                        row[3] = "PAPIER"
                    if row[3] in PNEUS:
                        row[3] = "PNEUS"
                    if row[3] in GRAVAT:
                        row[3] = "GRAVAT"
                    if row[3] in DIS:
                        row[3] = "DIS"
                    if row[3] in BOIS:
                        row[3] = "BOIS"
                    if row[3] in VEGETAUX:
                        row[3] = "VEGETAUX"
                    if row[3] in PLASTIQUE:
                        row[3] = "PLASTIQUE"
                    row = [row[2],row[3],row[6]]
                    writer.writerow(row)
            for row in readerPradie:
                if row:
                    if row[3] in FER:
                        row[3] = "FER"
                    if row[3] in DIB:
                        row[3] = "DIB"
                    if row[3] in DECHETS_ULTIMES:
                        row[3] = "DECHETS_ULTIMES"
                    if row[3] in CARTON:
                        row[3] = "CARTON"
                    if row[3] in DEEE:
                        row[3] = "DEEE"
                    if row[3] in PAPIER:
                        row[3] = "PAPIER"
                    if row[3] in PNEUS:
                        row[3] = "PNEUS"
                    if row[3] in GRAVAT:
                        row[3] = "GRAVAT"
                    if row[3] in DIS:
                        row[3] = "DIS"
                    if row[3] in BOIS:
                        row[3] = "BOIS"
                    if row[3] in VEGETAUX:
                        row[3] = "VEGETAUX"
                    if row[3] in PLASTIQUE:
                        row[3] = "PLASTIQUE"
                    row = [row[2],row[3],row[6]]
                    writer.writerow(row)
os.remove(filename)
os.remove(filename2)
filename = 'dechetTriesEntree.csv'

with open(filename, 'r') as csvfile:
    reader = csv.DictReader(csvfile)
    with open('DechetRegroupesEntrees.csv','w') as new_file:
        fieldnames = ['DateReception','TypeDNDEntre','Qte']
        writer = csv.DictWriter(new_file,fieldnames)
        writer.writeheader()
        def searchYear(year,typeDechet,kilo):
            d = next((item for item in dataFinal if item['DateReception'] == year), False)
            if d:
                for dechet in dataFinal:
                    if dechet['TypeDNDEntre'] == typeDechet and dechet['DateReception'] ==year:
                        dechet['Qte'] = int(float(dechet['Qte'])) + int(float(kilo))
                return False
            else:
                dataReturned = [
                    {"DateReception" : year, "TypeDNDEntre" : "FER", "Qte" :0},
                    {"DateReception" : year, "TypeDNDEntre" : "DIB", "Qte" :0},
                    {"DateReception" : year, "TypeDNDEntre" : "DECHETS_ULTIMES", "Qte" :0},
                    {"DateReception" : year, "TypeDNDEntre" : "CARTON", "Qte" :0},
                    {"DateReception" : year, "TypeDNDEntre" : "DEEE", "Qte" :0},
                    {"DateReception" : year, "TypeDNDEntre" : "PAPIER", "Qte" :0},
                    {"DateReception" : year, "TypeDNDEntre" : "PNEUS", "Qte" :0},
                    {"DateReception" : year, "TypeDNDEntre" : "GRAVAT", "Qte" :0},
                    {"DateReception" : year, "TypeDNDEntre" : "DIS", "Qte" :0},
                    {"DateReception" : year, "TypeDNDEntre" : "BOIS", "Qte" :0},
                    {"DateReception" : year, "TypeDNDEntre" : "VEGETAUX", "Qte" :0},
                    {"DateReception" : year, "TypeDNDEntre" : "PLASTIQUE", "Qte" :0}
                ]
                for findDechet in dataReturned:
                    if findDechet['TypeDNDEntre'] == typeDechet:
                        findDechet['Qte'] = int(float(kilo))
                    dataFinal.append(findDechet)
        dataFinal = []
        for row in reader:
            moisNumerique = row['DateReception'].split("/")[1]
            anneeNumerique = row['DateReception'].split("/")[2]
            tempsNumerique = moisNumerique + "/" + anneeNumerique
            if row:
                isDatePresent = searchYear(tempsNumerique, row['TypeDNDEntre'], row['Qte'])
        for row in dataFinal:
            writer.writerow(row)
os.remove(filename)
       
def csv_from_excel():
    wb = xlrd.open_workbook(fileSortie[0])
    print(wb.sheets())
    sh = wb.sheet_by_name('Sheet0')
    your_csv_file = open('sortie_bois_vert.csv', 'w')
    wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

    for rownum in range(sh.nrows):
        wr.writerow(sh.row_values(rownum))

    your_csv_file.close()

def csv_from_excel2():
    wb = xlrd.open_workbook(fileSortie[1])
    print(wb.sheets())
    sh = wb.sheet_by_name('Sheet0')
    your_csv_file = open('sortie_pradie.csv', 'w')
    wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

    for rownum in range(sh.nrows):
        wr.writerow(sh.row_values(rownum))

    your_csv_file.close()

# runs the csv_from_excel function:
csv_from_excel()
csv_from_excel2()


filename = 'sortie_bois_vert.csv'
filename2 = 'sortie_pradie.csv'

with open(filename2, 'r') as csvfile2:
    readerPradie = csv.reader(csvfile2)
    with open(filename, 'r') as csvfile:
        reader = csv.reader(csvfile)
        next(readerPradie)
        with open('dechetTriesSortie.csv','w+') as new_file:
            writer = csv.writer(new_file)
            FER = ["AGS BLANC ET PEINT","ALU COMPLEXE","ALU MELE","FERRAILLE","PLATIN","INOX","METAUX FERREUX","PLAQUES ALUMINIUM"]
            DIB = ["DECHETS","DECHETS VIDES","PLATRE","SELECTIF"]
            DECHETS_ULTIMES = ["DECHETS ULTIMES", "DECHETS VIDES", "DECHETS 191212"]
            CARTON = ["CARTON","SACS KRAFTS","JOURNAUX","ROGNURES DE CAISSERIE","CARTON A5","CARTON A4","AFFICHES","HUMIDITE"]
            DEEE = ["D3E","CUMULUS","NEONS"]
            PAPIER = ["AFNOR","PAPIER","BROCHURE","CHEQUES BROYES","ECRITS COULEURS","EXTRA CLAIR","BALLES DE LISTING","PAPIERS BROYES","ECRITS BLANCS"]
            PNEUS = ["PNEU","PNEUS"]
            GRAVAT = ["GRAVAT","GRAVATS VIDES"]
            BOIS = ["BOIS","BROYAT PALETTES","BROYAT DE BOIS B","BROYAT PALETTES AF","TERRE DE SOUCHE"]
            VEGETAUX = ["BROYAT DE VEGETAUX"]
            PLASTIQUE = ["PLASTIQUE","PARE CHOCS","BIG BAG","PLASTIQUE 95 5","PLASTIQUE 98 2","PLASTIQUES SANS VALEUR"]
            for row in reader:
                if row:
                    if row[1] in FER:
                        row[1] = "FER"
                    if row[1] in DIB:
                        row[1] = "DIB"
                    if row[1] in DECHETS_ULTIMES:
                        row[1] = "DECHETS_ULTIMES"
                    if row[1] in CARTON:
                        row[1] = "CARTON"
                    if row[1] in DEEE:
                        row[1] = "DEEE"
                    if row[1] in PAPIER:
                        row[1] = "PAPIER"
                    if row[1] in PNEUS:
                        row[1] = "PNEUS"
                    if row[1] in GRAVAT:
                        row[1] = "GRAVAT"
                    if row[1] in BOIS:
                        row[1] = "BOIS"
                    if row[1] in VEGETAUX:
                        row[1] = "VEGETAUX"
                    if row[1] in PLASTIQUE:
                        row[1] = "PLASTIQUE"
                    row = [row[0],row[1],row[3]]
                    writer.writerow(row)
            for row in readerPradie:
                if row:
                    if row[1] in FER:
                        row[1] = "FER"
                    if row[1] in DIB:
                        row[1] = "DIB"
                    if row[1] in DECHETS_ULTIMES:
                        row[1] = "DECHETS_ULTIMES"
                    if row[1] in CARTON:
                        row[1] = "CARTON"
                    if row[1] in DEEE:
                        row[1] = "DEEE"
                    if row[1] in PAPIER:
                        row[1] = "PAPIER"
                    if row[1] in PNEUS:
                        row[1] = "PNEUS"
                    if row[1] in GRAVAT:
                        row[1] = "GRAVAT"
                    if row[1] in BOIS:
                        row[1] = "BOIS"
                    if row[1] in VEGETAUX:
                        row[1] = "VEGETAUX"
                    if row[1] in PLASTIQUE:
                        row[1] = "PLASTIQUE"
                    row = [row[0],row[1],row[3]]
                    writer.writerow(row)
os.remove(filename)
os.remove(filename2)
filename = 'dechetTriesSortie.csv'

with open(filename, 'r') as csvfile:
    reader = csv.DictReader(csvfile)
    with open('DechetRegroupesSorties.csv','w') as new_file:
        fieldnames = ['DateExpe','DescriptionDechetExpedie','Qte']
        writer = csv.DictWriter(new_file,fieldnames)
        writer.writeheader()
        def searchYear(year,typeDechet,kilo):
            d = next((item for item in dataFinal if item['DateExpe'] == year), False)
            if d:
                for dechet in dataFinal:
                    if dechet['DescriptionDechetExpedie'] == typeDechet and dechet['DateExpe'] ==year:
                        dechet['Qte'] = int(float(dechet['Qte'])) + int(float(kilo))
                return False
            else:
                dataReturned = [
                    {"DateExpe" : year, "DescriptionDechetExpedie" : "FER", "Qte" :0},
                    {"DateExpe" : year, "DescriptionDechetExpedie" : "DIB", "Qte" :0},
                    {"DateExpe" : year, "DescriptionDechetExpedie" : "DECHETS_ULTIMES", "Qte" :0},
                    {"DateExpe" : year, "DescriptionDechetExpedie" : "CARTON", "Qte" :0},
                    {"DateExpe" : year, "DescriptionDechetExpedie" : "DEEE", "Qte" :0},
                    {"DateExpe" : year, "DescriptionDechetExpedie" : "PAPIER", "Qte" :0},
                    {"DateExpe" : year, "DescriptionDechetExpedie" : "PNEUS", "Qte" :0},
                    {"DateExpe" : year, "DescriptionDechetExpedie" : "GRAVAT", "Qte" :0},
                    {"DateExpe" : year, "DescriptionDechetExpedie" : "DIS", "Qte" :0},
                    {"DateExpe" : year, "DescriptionDechetExpedie" : "BOIS", "Qte" :0},
                    {"DateExpe" : year, "DescriptionDechetExpedie" : "VEGETAUX", "Qte" :0},
                    {"DateExpe" : year, "DescriptionDechetExpedie" : "PLASTIQUE", "Qte" :0}
                ]
                for findDechet in dataReturned:
                    if findDechet['DescriptionDechetExpedie'] == typeDechet:
                        findDechet['Qte'] = int(float(kilo))
                    dataFinal.append(findDechet)
        dataFinal = []
        for row in reader:
            moisNumerique = row['DateExpe'].split("/")[1]
            anneeNumerique = row['DateExpe'].split("/")[2]
            tempsNumerique = moisNumerique + "/" + anneeNumerique
            if row:
                isDatePresent = searchYear(tempsNumerique, row['DescriptionDechetExpedie'], row['Qte'])
        for row in dataFinal:
            writer.writerow(row)
os.remove(filename)
filename = 'DechetRegroupesEntrees.csv'
filename2 = 'DechetRegroupesSorties.csv'

with open(filename, 'r') as csvfile2:
    readerDechetEntrant = csv.DictReader(csvfile2)
    with open(filename2, 'r') as csvfile:
        readerDechetSortant = csv.DictReader(csvfile)
        with open('finalFile.csv','w+') as new_file:
            fieldnames = ['Date','TypeDechet','QteEntrant','QteSortant','Stock']
            writer = csv.DictWriter(new_file,fieldnames)
            writer.writeheader()
            for row in readerDechetEntrant:
                for line in readerDechetSortant:
                    if row['TypeDNDEntre'] == line['DescriptionDechetExpedie'] and row['DateReception'] == line['DateExpe']:
                        qteentrantTonne = 0
                        qteSortantTonne = 0
                        if int(row['Qte']) > 0:
                            qteentrantTonne = int(row['Qte']) / 1000
                        if int(line['Qte']) > 0:
                            qteSortantTonne = int(line['Qte']) / 1000
                        ligne = {'Date': row['DateReception'], 'TypeDechet': row['TypeDNDEntre'], 'QteEntrant': qteentrantTonne, 'QteSortant': qteSortantTonne, 'Stock' : 0}
                        writer.writerow(ligne)
                        break
os.remove(filename)
os.remove(filename2)

# Reading the csv file 
df_new = pd.read_csv('finalFile.csv') 
  
# saving xlsx file 
GFG = pd.ExcelWriter('entrees-sorties.xlsx') 
df_new.to_excel(GFG, index = False) 
  
GFG.save()
os.remove('finalFile.csv')
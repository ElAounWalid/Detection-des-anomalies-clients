import openpyxl
import random
from datetime import datetime, timedelta

#ouverture de fichier Excel 

wb = openpyxl.load_workbook('AGE.xlsx')

# Sélection de la feuille de calcul
sheet = wb['Age_F']


# Remplissage des cellules
for row_num in range(2,2002):
    start_date = datetime(1923, 2, 26)
    # Obtenir la date système
    end_date = datetime.now()

    #obtenir une date d'ouverture aleatoire
    # Calculer la différence entre les deux dates
    diff_dateDefaut_dateSys = end_date - start_date
    # Obtenir un nombre aléatoire entre 0 et le nombre de jours dans la différence de dates
    random_day = random.randint(0, diff_dateDefaut_dateSys.days)
    # Ajouter le nombre de jours aléatoire à la date de début pour obtenir la date aléatoire
    date_naissance = start_date + timedelta(days=random_day)
    
    sheet.cell(row=row_num, column=11).value = date_naissance


# Sauvegarde du fichier
wb.save('AGE.xlsx')
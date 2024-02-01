import openpyxl
import random


def generate_CIN_number():
    """Génère un numéro de CIN aléatoire"""
    CIN_number = str(random.randint(0, 1))
    for i in range(7):
        CIN_number += str(random.randint(0, 9))  # Ajoute un chiffre aléatoire entre 0 et 9
    return CIN_number


#ouverture de fichier Excel 
wb = openpyxl.load_workbook('CIN.xlsx')

# Sélection de la feuille de calcul
sheet = wb['CIN_F']

# Remplissage des cellules
for row_num in range(2,2002):
    sheet.cell(row=row_num, column=11).value = generate_CIN_number()

# Sauvegarde du fichier
wb.save('CIN.xlsx')


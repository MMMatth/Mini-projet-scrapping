## import the module ##
from openpyxl import load_workbook, Workbook
import os , requests, shutil

## function ##
# Fonction qui permet de modifier une case d'un fichier excel
def modif_cellule(fichier, feuille, ligne, colonne, valeur):
    """
    modif_cellule est une fonction qui permet de modifier une case d'un fichier excel

    Args:
        fichier (str): lien du fichier excel
        feuille (str): nom de la feuille
        ligne (int): numéro de la ligne
        colonne (int): numéro de la colonne
        valeur (_type_): valeur à mettre dans la case
    """
    wb = load_workbook(fichier)
    ws = wb[feuille]
    ws.cell(row=ligne, column=colonne).value = valeur
    wb.save(fichier)

## class ##
class pdf:
    def __init__(self, lien : str, nom : str):
        self.nom = nom # nom du prof
        self.lien = lien  # lien du site du prof
        self.page = requests.get(self.lien)   # page du site du prof              
        self.nbr_dl = 0 # nombre de pdf téléchargé
        self.pdf = []  # liste des liens pdf
        self.recup_pdf()
    
    def __str__(self):
        """__str__ permet de définir ce que l'on veut afficher quand on affiche l'objet"""
        return self.lien
    
    def recup_pdf(self):
        """recup_pdf permet de récupérer les liens pdf sur le site du prof"""
        for i in range(len(self.page.text)):
            if self.page.text[i:i+4] == ".pdf": # si on trouve ".pdf"
                for j in range (i, 0, -1): # on remonte dans la page jusqu'à trouver le début du lien
                    if self.page.text[j] == '"' or self.page.text[j] == "'": # si on trouve le début du lien
                        self.pdf.append(self.page.text[j+1:i+4]) # on ajoute le lien à la liste
                        break
    
    def verif_url(self, url : str):
        """verif_url permet de verifier si le lien est bien une url et pas un lien relatif"""
        return url[0:4] == "http"

    
    def dl_pdf(self, path : str):
        """
        dl_pdf est une fonction qui permet de télécharger les pdf dans un dossier

        Args:
            path (str): chemin du dossier où télécharger les pdf
        """
        for i_lien in range(len(self.pdf)): 
            if self.verif_url(self.pdf[i_lien]): # si c'est une url et pas un lien relatif
                r = requests.get(self.pdf[i_lien], stream=True) # on récupère le pdf
                with open(path + str(i_lien) + ".pdf", 'wb') as f: # on le télécharge dans le dossier
                    f.write(r.content) # on écrit le contenu du pdf dans le fichier 
                self.nbr_dl += 1 # on incrémente le nombre de pdf téléchargé
                
    
      
         

## main ##
## lire le fichier excel et mettre les données dans un dictionnaire ##
wb = load_workbook(filename = 'liste_ens.xlsx')
sheet_ranges = wb['Feuil1']
site = {
    sheet_ranges['B2'].value: [sheet_ranges['C2'].value, None],
    sheet_ranges['B3'].value: [sheet_ranges['C3'].value, None],
    sheet_ranges['B4'].value: [sheet_ranges['C4'].value, None],
}

# ## on suprime les anciens fichiers ##
# for prenom in site.keys():
#     shutil.rmtree(prenom)

## crée un dossier par professeur ##
for prenom in site.keys(): 
    os.mkdir(prenom)
    
## on récupère les liens des pdf sur les sites ##
l_table = 1 
for prenom in site.keys():
    site[prenom][1] = pdf(site[prenom][0],prenom) # on crée un objet pdf pour chaque prof
    site[prenom][1].dl_pdf(str(prenom)+"/"+str(prenom)) # on télécharge les pdf dans le dossier de chaque prof
    modif_cellule("liste_ens.xlsx", "Feuil1", l_table + 1, 4, site[prenom][1].nbr_dl) # on modifie le fichier excel
    l_table += 1 # on incrémente l_table pour aller à la ligne suivante du fichier excel
    

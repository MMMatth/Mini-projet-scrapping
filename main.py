## import the module ##
from openpyxl import load_workbook, Workbook
import os , requests, shutil

## function ##
# Fontion qui recupère les liens pdf sur le site
class pdf:
    def __init__(self, lien : str, nom : str):
        self.nom = nom # nom du prof
        self.lien = lien  # lien du site du prof
        self.page = requests.get(self.lien)   # page du site du prof              
        self.nbr_dl = 0 # nombre de pdf téléchargé
        self.pdf = []  # liste des liens pdf
        self.recup_pdf()
    
    def __str__(self):
        """
        __str__ permet de définir ce que l'on veut afficher quand on affiche l'objet

        Returns:
            str : le lien du site du prof
        """
        return self.lien
    
    def recup_pdf(self):
        """
        recup_pdf permet de récupérer les liens pdf sur le site du prof
        """
        for i in range(len(self.page.text)):
            if self.page.text[i:i+4] == ".pdf": # si on trouve ".pdf"
                for j in range (i, 0, -1): # on remonte dans la page jusqu'à trouver le début du lien
                    if self.page.text[j] == '"' or self.page.text[j] == "'": # si on trouve le début du lien
                        self.pdf.append(self.page.text[j+1:i+4]) # on ajoute le lien à la liste
                        break
    
    def verif_url(self, url):
        """
        verif_url permet de verifier si le lien est bien une url et pas un lien relatif

        Args:
            url (str): lien à vérifier

        Returns:
            booléen : True si c'est une url, False sinon
        """
        if url[0:4] != "http":
            return False
        return True
        
    
    def dl_pdf(self, path):
        for i_lien in range(2):
            if self.verif_url(self.pdf[i_lien]):
                r = requests.get(self.pdf[i_lien], stream=True)
                with open(path + str(i_lien) + ".pdf", 'wb') as f:
                    f.write(r.content)
                self.nbr_dl += 1
                
    
      
         

## main ##

## lire le fichier excel et mettre les données dans un dictionnaire ##
wb = load_workbook(filename = 'liste_ens.xlsx')
sheet_ranges = wb['Feuil1']
site = {
    sheet_ranges['B2'].value: [sheet_ranges['C2'].value, None],
    sheet_ranges['B3'].value: [sheet_ranges['C3'].value, None],
    sheet_ranges['B4'].value: [sheet_ranges['C4'].value, None],
}

## on suprime les anciens fichiers ##
for prenom in site.keys():
    shutil.rmtree(prenom)

## crée un dossier par professeur ##
for prenom in site.keys():
    os.mkdir(prenom)
    
## on récupère les liens des pdf sur les sites ##
indice = 1
for prenom in site.keys():
    site[prenom][1] = pdf(site[prenom][0],prenom)
    site[prenom][1].dl_pdf(str(prenom)+"/"+str(prenom))
    # modifie le fichier excel liste_ens.xlsx
    wb = load_workbook(filename = 'liste_ens.xlsx')
    sheet_ranges = wb['Feuil1']
    sheet_ranges.cell("D"+str(indice)).value = site[prenom][1].nbr_dl
    indice += 1
    

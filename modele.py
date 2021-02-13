import requests # requests
import csv # csv
from datetime import datetime
import xlsxwriter # xlsxwriter
from pathlib import Path



class fichierExcel:

    """ Constructeur""" 
    def __init__(self):
        self.nb_erreur_csv = 0
        self.nb_erreur_xlsx = 0
        self.lignes_excel = []

    """ méthode qui permet l'ajout d'une ligne dans notre fichier"""
    def ajouter_ligne(self,ligne):
        self.lignes_excel.append(ligne)

    """ méthode qui permet d'écrire le fichier au format csv"""
    def toCSV(self):
        # utilisation de la librairie datetime afin de générer le nom du fichier
        now = datetime.now()
        dt_string = now.strftime("%d_%m_%Y_%H_%M_%S")
        d_string = now.strftime("%d_%m_%Y")
        # Creation de repertoire s'il n'existe pas
        path = "data/"+d_string+"/CSV/"
        Path(path).mkdir(parents=True, exist_ok=True)
        try:
            with open(path+'entreprise'+dt_string+'.csv', mode='w',encoding="latin-1") as entreprise_file:
                entreprise_writer = csv.writer(entreprise_file, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                # entête
                entreprise_writer.writerow(["Nom",'Siège CP','Siège Ville',"activite","Tranche d'effectif","Dirigeant","Forme juridique",'idcc'])
                # Corps
                for ligne in self.lignes_excel:
                    try:
                        entreprise_writer.writerow([
                            ligne.nom,
                            ligne.zip,
                            ligne.ville,
                            ligne.activite,
                            ligne.tranche_effectif,
                            ligne.dirigeant,
                            ligne.forme_juridique,
                            ligne.idcc,
                        ])
                    except:
                        self.nb_erreur_csv +=1
                        pass # Si une ligne cause une exception on passe à la ligne suivante
            print("Sauvegarde réussie dans fichier"+dt_string+".csv")
        except:
            print(" Il y a eu une erreur lors de la sauvegarde du fichier csv")
        

    """ méthode qui permet d'écrire le fichier au format xlsl"""
    def toxlsx(self):
            # utilisation de la librairie datetime afin de générer le nom du fichier
            now = datetime.now()
            dt_string = now.strftime("%d_%m_%Y_%H_%M_%S")
            d_string = now.strftime("%d_%m_%Y")
            # Create a workbook and add a worksheet.
            path = "data/"+d_string+"/XLXS/"
            Path(path).mkdir(parents=True, exist_ok=True)
            workbook = xlsxwriter.Workbook(path+'entreprise-'+dt_string+'.xlsx')
            worksheet = workbook.add_worksheet()
            style = workbook.add_format({'bold': True,'bg_color':'green'})
            worksheet.write_row(0,0,["Nom",'Siège CP','Siège Ville',"Activite","Tranche d'effectif","Dirigeant","Forme juridique",'IDCC'],style)
            # Iterate over the data and write it out row by row.
            row = 1
            for ligne in self.lignes_excel:
                try:
                    worksheet.write_row(row,0,[ligne.nom,
                                ligne.zip,
                                ligne.ville,
                                ligne.activite,
                                ligne.tranche_effectif,
                                str(ligne.dirigeant),
                                ligne.forme_juridique,
                                str(ligne.idcc)])
                    row += 1
                except:
                    self.nb_erreur_xlsx +=1
                    pass
                    

            workbook.close()
        


class ligneExcel:


    """
        Constructeur
        @param adresse => adresse de l'entreprise
        @param activite => activite de l'entreprise
        @param effectif => effectif de l'entreprise
        @param dirigeant => dirigeant de l'entreprise
        @param forme_juridique => forme_juridique de l'entreprise
        @param idcc => idcc de l'entreprise
    """
    def __init__(self,nom,ville,zip,activite,tranche_effectif,dirigeant,forme_juridique,idcc):
        try:
            self.nom = nom
            self.ville = ville
            self.zip = zip
            self.activite = activite 
            self.tranche_effectif = tranche_effectif
            self.dirigeant = dirigeant
            self.forme_juridique = forme_juridique
            self.idcc = idcc
        except:
            pass

class Requete:
    def __init__(self):
        self.params = {
            "par_page":10000,
            "entreprise_cessee":"false"
        }
        self.requete ="https://api.pappers.fr/v1/recherche"
        self.response = None
        self.fichier =  fichierExcel() # Création d'un fichier excel

    def recuperer_token(self):
        try:
            with open("token.txt", "r") as f:
                self.params["api_token"] = f.read()
                return 1
        except: 
            print("Impossible de lire le fichier Token")
            return 0
            

    def demander_effectif(self,type):
        effectif=input("Veuillez saisir la tranche d'effectif "+str(type)+ " au format sirène (XX) : ")
        while len(effectif) != 2 or not effectif.isdigit():
            print("Format incorrect")
            effectif=input("Veuillez saisir la tranche d'effectif "+str(type)+" au format sirène (XX) : ")
        self.params["tranche_effectif_"+str(type)] = effectif
    
    def demander_convention_collective(self):
        convention_collective=input("Veuillez saisir le code d'indentification de convention collective (IDCC) (XXXX) : ")
        while len(convention_collective) != 4 or not convention_collective.isdigit():
            print("Format incorrect")
            convention_collective=input("Veuillez saisir le code d'indentification de convention collective (IDCC) (XXXX) : ")
        self.params["convention_collective"] = convention_collective

    
    def executer_requete(self):
        self.demander_effectif('min')
        self.demander_effectif('max')
        self.demander_convention_collective()
        if self.recuperer_token() :
            self.response = requests.get(self.requete,params=self.params)
            self.analyse_code_retour()
    
    def analyse_code_retour(self):
        if self.response.status_code == 200:
            self.traitement_requete()
        elif self.response.status_code == 400:
            print("400 : Paramètre de la requête incorrects")
            return 0
        elif self.response.status_code == 401:
            print("401 : Clé API incorrecte")
            return 0
        elif self.response.status_code == 404:
            print("404 : Document indisponible")
            return 0
        elif self.response.status_code == 503:
            print("503 : Service indisponible")
            return 0
        

    def traitement_requete(self):
        for entreprise in self.response.json().get('entreprises'):
            # Etape 1 => map => on récupère uniquement l'idcc
            # Etape 2 => set => on garde une clef unique
            idcc = set(map(lambda conv: conv.get('idcc'),entreprise.get('conventions_collectives')))
            ligne = ligneExcel(
                    entreprise.get('nom_entreprise'),
                    entreprise.get('siege').get("ville"),
                    entreprise.get('siege').get("code_postal"),
                    entreprise.get("domaine_activite"),
                    entreprise.get("effectif"),
                    list(filter(lambda representant: representant.get('qualite') == "Gérant",entreprise.get("representants"))) if entreprise.get('representants') else None,
                    entreprise.get("forme_juridique"),
                   
                    idcc if idcc else None)
            self.fichier.ajouter_ligne(ligne)
        self.fichier.toCSV()
        self.fichier.toxlsx()
        self.afficher_resultat()

    def afficher_resultat(self):
        print("______________________")
        print("Nombre total d'entreprises : "+str(self.response.json().get('total_entreprises')))
        print("Entreprises non sauvegardées dans le fichier CSV : "+str(self.fichier.nb_erreur_csv))
        print("Entreprises sauvegardées dans le fichier CSV : "+str(self.response.json().get('total_entreprises') - self.fichier.nb_erreur_csv))
        print("________")
        print("Entreprises non sauvegardées dans le fichier XLXS : "+str(self.fichier.nb_erreur_xlsx))
        print("Entreprises sauvegardées dans le fichier XLXS : "+str(self.response.json().get('total_entreprises') - self.fichier.nb_erreur_xlsx))
        


r = Requete()
r.executer_requete()



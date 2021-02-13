# Matchers_Finder

Installer les bibliothèques "request", "xlsxwriter", "csv", "pathlib" (à installer si non présents dans l'OS) pour python
Le programme s'exécute via l'invite de commande grâce à la commande "python3 nom_du_code.py" (sous linux)
 Le code token pourrait être stocké dans un autre fichier pour plus de sécurité.

Pour l'effectif de l'entreprise, entrer les limites au format sirène (https://www.sirene.fr/sirene/public/variable/tefen#:~:text=Cette%20variable%20correspond%20%C3%A0%20la,effectif%20salari%C3%A9%20de%20l'entreprise.)
ex: dans notre cas, on cherche les entreprises dont l'effectif est compris entre 01 et 49 salariés, donc en effectif min on met "01" et en max on met "12" (cf. lien ci-dessus).
Les fichiers CSV et XLSX seront créés dans un dossier "./data/date_du_jour/extension_du_fichier/nom_du_fichier" 

En utilisant la classe "ligneExcel" on peut y ajouter un autre attribut de classe qui correspondrait au numéro de téléphone de l'entreprise, qu'on peut obtenir via l'API des pages jaunes par exemple.

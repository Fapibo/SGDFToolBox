# But du programme
A partir de l'extraction des individus de l'intranet scout, ce programme:
- génère un fichier CSV pour importer les contacts du groupe dans gmail
- génère un excel de synthèse des membres du groupe


# Mode d'emploi
## Extraction des données depuis l'intranet SGDF
Dans le menu de l'intranet: 
`pilotage > Extractions > Extraire individus`

Cocher: 
* Fonctions principales et secondaires
* Coordonnées des individus
* Coordonnées des Parents
* Inscription

*Ne pas cocher adhésion* sinon, on n'a pas les invités et les inscriptions non terminées!

Laisser toutes les autres options par défaut. Il est possible d'enregister la configuration.

## utilisation
1. Lancer le programe
1. Selectionner le fichier extrait de l'intranet
1. Définir les options. <br> Explication de l'option **Ne pas limiter les doublons de mails**: si l'intranet contient le même mail pour le jeune, le papa et la maman, on aura trois contacts avec le même mail. Si décoché, seul le mail de la maman sera exporté.
1. cliquer sur __"Go"__


Le programme crée dans le même dossier:
* un fichier CSV à importer dans gmail
* un fichier excel

# Lancement
* __Windows uniquement__: Lancement de l'executable avec un double clic
Lancer l'executable en ignorant les avertissements de sécurité de windows

note: L'executable est généré par pyinstaller: `pyinstaller SGDF_Toolbox.py`

* __tout système__: directement en python
1. installer python 3
1. installer les dépendances avec la commande
`pip3 install -r requirements.txt`
1. lancer le script

# Fichier de configuration "Const.py"
Le fichier de configuration Const.py n'a normalement pas besoin d'être modifié par l'utilisateur.
Il contient (voir les explications directement dans le fichier):
- l'ordre des colonnes dans le fichier importé
- l'ordre des colonnes dans le fichier CSV à générer
- les types de structure à chercher ainsi que leur nom court.  

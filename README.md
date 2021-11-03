# Intro
- génère les fichier CSV pour importer les contacts dans gmail
- génère un excel de contacts des jeunes et des parents (+vérif des dates de naissance)


# Mode d'emplo
## Extraction des données depuis l'intranet SGDF
Dans l'intranet: pilatege>extraction>extraire individus
Cocher: 
* Fonctions principales et secondaires
*Coordonnées des individus
*Coordonnées des Parents
*Inscription

Ne pas cocher adhésion sinon, on n'a pas les invités et les inscriptions non terminées!

laisser toutes les autres options par défaut

## utilisation
1. lancer le programe
1. selectionner le fichier extrait de l'intranet
1. définir les options:
- inclure les pré-inscrits
- inclure les invités
- Ne pas limiter les doublons de mails: si l'intranet contient le même mail pour le jeune, le papa et la maman, on aura 3 contacts avec le même mail. Si décoché, seul le mail de la maman sera exporté.
1. cliquer sur "Go"

Le programme crée dans le même dossier:
* un fichier CSV à importer dans gmail
* un fichier excel

# Lancement
* Windows uniquement: Utilisation de l'executable windows
Lancer l'executable en ignorant les avertissements de sécurité de windows

note: L'executable est généré par pyinstaller: `pyinstaller SGDF_Toolbox.py`

* tout système: directement en python
1. installer python 3
1. installer les dépendances avec la commande
`pip3 install -r requirements.txt`
1. lancer le script

## Fichier de configuration "Const.py"
Le fichier de configuration Const.py n'a normalement pas besoin d'être modifié par l'utilisateur.
Il contient (voir commentaires dans le fichier):
- l'ordre des colonnes dans le fichier importé
- l'ordre des colonnes dans le fichier CSV à générer
- les types de structure à chercher ainsi que leur nom court.  

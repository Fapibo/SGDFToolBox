########### FICHIER des adherents ###############
# mettre le numero de la colonne EN COMMENCEANT PAR 0
# cocher individu paretns et inscription mais pas adhesion!
# 1ere colonne = colonne 0
ColCiv = 0 # IndividuCivilite.NomCourt
ColNom = 2 # Individu.Nom
ColPrenom = 3 # Individu.Prenom
Colstruct = 5 # Structure.Nom
ColFonction = 6 # Fonction.Code
ColMaitrise = 7 # Fonction.CategorieMembre
ColFoncSecondaire = 10 # Inscription.Delegations
ColAdr1 = 12 # Individu.Adresse.Ligne1
ColAdr2 = 13 # Individu.Adresse.Ligne2
ColAdr3 = 14 # Individu.Adresse.Ligne3
ColCP = 15 # Individu.Adresse.CodePostal
ColVille = 16 # Individu.Adresse.Municipalite
ColPays = 17 # Individu.Adresse.Pays
ColTel1 = 18 # Individu.TelephoneDomicile
ColTel2 = 19 # Individu.TelephonePortable1
ColTel3 = 20 # Individu.TelephonePortable2
ColTel4 = 21 # Individu.TelephoneBureau
ColMail1 = 23 # Individu.CourrielPersonnel
ColMail2 = 24 # Individu.CourrielDédiéSGDF
ColDateN = 25 # Date de naissance
ColPapaNom = 42 # Pere.Nom
ColPapaPrenom = 43 # Pere.Prenom
ColPapaAdr1 = 44 # Pere.Adresse.Ligne1
ColPapaAdr2 = 45 # Pere.Adresse.Ligne2
ColPapaAdr3 = 46 # Pere.Adresse.Ligne3
ColPapaCP = 47 # Pere.Adresse.CodePostal
ColPapaVille = 48 # Pere.Adresse.Municipalite
ColPapaPays = 49 # Pere.Adresse.Pays
ColPapaTel1 = 50 # Pere.TelephoneDomicile
ColPapaTel2 = 51 # Pere.TelephonePortable1
ColPapaTel3 = 52 # Pere.TelephonePortable2
ColPapaTel4 = 53 # Pere.TelephoneBureau
ColPapaMail1 = 55 # Pere.CourrielPersonnel
ColPapaMail2 = 56 # Pere.CourrielDédiéSGDF
ColMamNom = 59 # Mere.Nom
ColMamPrenom = 60 # Mere.Prenom
ColMamAdr1 = 61 # Mere.Adresse.Ligne1
ColMamAdr2 = 62 # Mere.Adresse.Ligne2
ColMamAdr3 = 63 # Mere.Adresse.Ligne3
ColMamCP = 64 # Mere.Adresse.CodePostal
ColMamVille = 65 # Mere.Adresse.Municipalite
ColMamPays = 66 # Mere.Adresse.Pays
ColMamTel1 = 67 # Mere.TelephoneDomicile
ColMamTel2 = 68 # Mere.TelephonePortable1
ColMamTel3 = 69 # Mere.TelephonePortable2
ColMamTel4 = 70 # Mere.TelephoneBureau
ColMamMail1 = 72 # Mere.CourrielPersonnel
ColMamMail2 = 73 # Mere.CourrielDédiéSGDF
ColInscrit = 74 # Inscriptions.Type
ColInscDateFin = 76 # Inscriptions.Date.fin

###########
TxtPreInscrit = "Pre-inscrit"
TxtInvit = "Invit"
TxtInscrit = "Inscrit"
TxT_LJ = 'LOUVETEAU' # texte forcement contenu dans Structure.Nom
TxT_SG = 'SCOUTS GUIDES' # texte forcement contenu dans Structure.Nom
TxT_PK = 'PIONNIERS CARAVELLES' # texte forcement contenu dans Structure.Nom
TxT_Farfa = 'FARFADETS' # texte forcement contenu dans Structure.Nom
Txt_Violets = 'GROUPE DE VAISE'
TxtCompa = 'COMPAGNONS - VAISE'
Label_LJ = 'Oranges' 
Label_SG = 'Bleus' 
Label_PK = 'Rouges' 
Label_Farfa = 'Farfas' 
Label_Violets = 'Violets'
LabelCompa = 'Compas'
AgeMinLJ = 7
AgeMinSG = 10
AgeMinPK = 13
AgeMinComp = 16
# cette liste de dictionnaire contient les infos de la structure
# Label = nom cours
# Txt = Texte a chercher dans les données de l'intranet pour trouver l'unité (s'il y a plusieurs unités, ne pas mettre le numéro)
# age min sert pour vérifier si l'enfant est bien inscrit dans la bonne couleur
Struct_Type = [
                {"Label": 'Oranges' , "Txt":'LOUVETEAU' ,  "AgeMin": 8, "AgeMax": 10},
                {"Label": 'Bleus' , "Txt":'SCOUTS GUIDES' ,  "AgeMin": 11, "AgeMax": 13},
                {"Label": 'Rouges' , "Txt":'PIONNIERS CARAVELLES' ,  "AgeMin": 14, "AgeMax": 16},
                {"Label": 'Farfas' , "Txt":'FARFADETS' ,  "AgeMin": 6, "AgeMax": 7},
                {"Label": 'Violets' , "Txt":'GROUPE DE ' ,  "AgeMin": 17, "AgeMax": 999},
                {"Label": 'Compas' , "Txt":'COMPAGNONS - ' ,  "AgeMin": 17, "AgeMax": 999}
            ]



# Format contact
#* My Contacts ::: majyyyy10dd-162508 ::: Parents ::: Parents rouge

###########
CSVNbCol = 60
CSV_Header = 'Name,Given Name,Additional Name,Family Name,Yomi Name,Given Name Yomi,Additional Name Yomi,Family Name Yomi,Name Prefix,Name Suffix,Initials,Nickname,Short Name,Maiden Name,Birthday,Gender,Location,Billing Information,Directory Server,Mileage,Occupation,Hobby,Sensitivity,Priority,Subject,Notes,Language,Photo,Group Membership,E-mail 1 - Type,E-mail 1 - Value,E-mail 2 - Type,E-mail 2 - Value,E-mail 3 - Type,E-mail 3 - Value,Phone 1 - Type,Phone 1 - Value,Phone 2 - Type,Phone 2 - Value,Phone 3 - Type,Phone 3 - Value,Phone 4 - Type,Phone 4 - Value,Address 1 - Type,Address 1 - Formatted,Address 1 - Street,Address 1 - City,Address 1 - PO Box,Address 1 - Region,Address 1 - Postal Code,Address 1 - Country,Address 1 - Extended Address,Organization 1 - Type,Organization 1 - Name,Organization 1 - Yomi Name,Organization 1 - Title,Organization 1 - Department,Organization 1 - Symbol,Organization 1 - Location,Organization 1 - Job Description'
CSV_NomLong = 0 # Name
CSV_Prenom = 1 # Given Name
CSV_Nom = 3 # Family Name
CSV_Prefixe = 8 # Name Prefix
CSV_Suffixe = 9 # Name Suffix
CSV_notes = 25 # Notes: invite ou inscrit ou pre-inscrit
CSV_Group = 28 # Group Membership --> etiquette, séparée par ' ::: '
CSV_mail = 30 # E-mail 1 - Value
CSV_tel1 = 36 # Phone 1 - Value
CSV_tel2 = 38 # Phone 2 - Value
CSV_tel3 = 40 # Phone 3 - Value
CSV_tel4 = 42 # Phone 4 - Value
CSV_TelType1 = 35 # Phone 1 - Type
CSV_TelType2 = 37 # Phone 2 - Type
CSV_TelType3 = 39 # Phone 3 - Type
CSV_TelType4 = 41 # Phone 4 - Type
CSV_adr_full = 44 # Address 1 - Formatted
CSV_Adr1 = 45 # Address 1 - Street
CSV_Adr2 = 51 # Address 1 - Extended Address
CSV_ville = 46 # Address 1 - City
CSV_CP = 49 # Address 1 - Postal Code
CSV_Pays = 50 # Address 1 - Country
CSV_Structure = 53 # Organization 1 - Name

# Misc
URLHelp = 'https://github.com/Fapibo/SGDFToolBox#readme'


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
ColPrenomCivil = 12 #Individu.PrenomCivil
ColAdr1 = 13 # Individu.Adresse.Ligne1
ColAdr2 = 14 # Individu.Adresse.Ligne2
ColAdr3 = 15 # Individu.Adresse.Ligne3
ColCP = 16 # Individu.Adresse.CodePostal
ColVille = 17 # Individu.Adresse.Municipalite
ColPays = 18 # Individu.Adresse.Pays
ColTel1 = 19 # Individu.TelephoneDomicile
ColTel2 = 20 # Individu.TelephonePortable1
ColTel3 = 21 # Individu.TelephonePortable2
ColTel4 = 22 # Individu.TelephoneBureau
ColMail1 = 24 # Individu.CourrielPersonnel
ColMail2 = 25 # Individu.CourrielDédiéSGDF
ColDateN = 26 # Date de naissance
ColNumAlloc = 30 # Individu.NumeroAllocataire
ColDroitsImage = 36#Individu.DroitImageExterne
ColRespL1Civ = 43 # PereCivilite.NomCourt
ColRespL1Nom = 45 # Pere.Nom
ColRespL1Prenom = 46 # Pere.Prenom
ColRespL1Adr1 = 47 # Pere.Adresse.Ligne1
ColRespL1Adr2 = 48 # Pere.Adresse.Ligne2
ColRespL1Adr3 = 49 # Pere.Adresse.Ligne3
ColRespL1CP = 50 # Pere.Adresse.CodePostal
ColRespL1Ville = 51 # Pere.Adresse.Municipalite
ColRespL1Pays = 52 # Pere.Adresse.Pays
ColRespL1Tel1 = 53 # Pere.TelephoneDomicile
ColRespL1Tel2 = 54 # Pere.TelephonePortable1
ColRespL1Tel3 = 55 # Pere.TelephonePortable2
ColRespL1Tel4 = 56 # Pere.TelephoneBureau
ColRespL1Mail1 = 58 # Pere.CourrielPersonnel
ColRespL1Mail2 = 59 # Pere.CourrielDédiéSGDF
ColRespL2Civ = 60 # MereCivilite.NomCourt
ColRespL2Nom = 62 # Mere.Nom
ColRespL2Prenom = 63 # Mere.Prenom
ColRespL2Adr1 = 64 # Mere.Adresse.Ligne1
ColRespL2Adr2 = 65 # Mere.Adresse.Ligne2
ColRespL2Adr3 = 66 # Mere.Adresse.Ligne3
ColRespL2CP = 67 # Mere.Adresse.CodePostal
ColRespL2Ville = 68 # Mere.Adresse.Municipalite
ColRespL2Pays = 69 # Mere.Adresse.Pays
ColRespL2Tel1 = 70 # Mere.TelephoneDomicile
ColRespL2Tel2 = 71 # Mere.TelephonePortable1
ColRespL2Tel3 = 72 # Mere.TelephonePortale2
ColRespL2Tel4 = 73 # Mere.TelephoneBureau
ColRespL2Mail1 = 75 # Mere.CourrielPersonnel
ColRespL2Mail2 = 76 # Mere.CourrielDédiéSGDF
ColInscrit = 94 # Inscriptions.Type
ColInscDateDebut = 95 # Inscriptions.Date.Debut
ColInscDateFin = 96 # Inscriptions.Date.fin

###########
TxtPreInscrit = "-inscrit"
TxtInvit = "Invit"
TxtInscrit = "Inscrit"
TxtFille = 'e' # texte contenu dans la colonne civité qui permet de savoir si c'est une fille -- lettre "e"
LabelMembre = 'Membre(s)'
LabelParent = 'Parent(s)'
LabelMaitrise = 'Maitrise'
LabelRespL1 = 'Resp1'
LabelRespL2 = 'Resp2'
LabelTousLesParents = 'TOUS_les_parents'
LabelTouslesMembres = 'TOUS_les_membres'
LabelChefs = 'TOUS_les_Chefs LJ SG PK'

# cette liste de dictionnaire contient les infos de la structure
# Label = nom cours
# Txt = Texte a chercher dans les données de l'intranet pour trouver l'unité (s'il y a plusieurs unités, ne pas mettre le numéro)
# age min sert pour vérifier si l'enfant est bien inscrit dans la bonne couleur

# utile pour detecter les compas en fonciton secondaires et pour ajouter les maitrises Farfa et Compas aux violets
TxtCompaFnSecondr = '- 1ERE COMPAGNONS - VAISE'
LabelCompa = 'Compas'

Struct_Type = [
                {"Label": 'Oranges' , "Txt":'LOUVETEAU' ,  "AgeMin": 8, "AgeMax": 10},
                {"Label": 'Bleus' , "Txt":'SCOUTS GUIDES' ,  "AgeMin": 11, "AgeMax": 13},
                {"Label": 'Rouges' , "Txt":'PIONNIERS CARAVELLES' ,  "AgeMin": 14, "AgeMax": 16},
                {"Label": 'Farfas', "Txt":'FARFADETS' ,  "AgeMin": 6, "AgeMax": 7},
                {"Label": 'Violets' , "Txt":'GROUPE DE ' ,  "AgeMin": 17, "AgeMax": 99},
                {"Label": 'Impeesa' , "Txt":'IMPEESA' ,  "AgeMin": 17, "AgeMax": 99},
                {"Label": LabelCompa , "Txt": 'COMPAGNONS - ' ,  "AgeMin": 17, "AgeMax": 99}
            ]

SelectLabelChefs = Struct_Type[0]['Label'] + Struct_Type[1]['Label'] + Struct_Type[2]['Label']

########### pour génération du CSV pour GMAIL
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



# https://github.com/Fapibo/SGDFToolBox
from tkinter import filedialog, Tk, StringVar, IntVar, Label, Button, Checkbutton
import datetime
import webbrowser
import pandas as pd
import numpy as np
from styleframe  import StyleFrame
import os
import csv
import re
import xlsxwriter
from Const import * # import constant file

class SGDF_ToolboxUI:
    def __init__(self, master):
        pd.options.display.max_rows = 12
        self.master = master
        self.table = [] # fichier source importé sous forme de liste 2D
        self.tableOK = False # True quand fichier importé OK.

        self.ContactsDic = {} # contient toutes les contacts et leurs infos utiles sous forme { nom: [nom1 nom2 nom3 ...] , prenom : [prenom1 prenom2 prenom3], structure: [.....) }
        self.ContactsDF = pd.DataFrame() # dataframe  issue du dictionnaire ci dessus pour pouvoir manipuler les données
        self.Structures =  {}  # Nom vers alias des unités du Groupe + année naiss min et max. Ex: {'1ERE LOUVETEAUX JEANNETTES - VAISE': [ 'Orange_1' 2011 2013 ]}
        # IHM
        master.title("SGDF_Toolbox")
        master.minsize(300, 50)
        
        # Options
        self.SourceFile = StringVar()
        self.Info = StringVar()
        self.Incl_Pre_inscrits = IntVar()
        self.Incl_Invit = IntVar()
        self.Incl_Invit.set(1)
        self.KeepDuplicatesMails = IntVar()
       
        self.Infolabel = Label(master, textvariable=self.Info)
        self.Infolabel.grid(row=0, column=1)
        self.InfoMsg("\"Parcourir...\" pour choisir un fichier.", "Info")

        self.Captionlabel = Label(master, text="Fichier source:")
        self.Captionlabel.grid(row=1, column=0)
        
        self.HelpLink = Label(master, text="Lien vers aide et code source", fg="blue", cursor="hand2")
        self.HelpLink.grid(row=0, column=0)
        self.HelpLink.bind("<Button-1>", lambda e: webbrowser.open_new(URLHelp))

        self.buttonBrowse = Button(master, text="Parcourir...", command=self.BBrowse)
        self.buttonBrowse.grid(row=1, column=2)
 
        self.Pathlabel = Label(master, textvariable=self.SourceFile)
        self.Pathlabel.grid(row=1, column=1)
        
        self.Go_button = Button(master, text="Go!", command=lambda: self.BGo())
        self.Go_button.grid(row=3, column=1)
        
        self.CheckPreInscrit = Checkbutton(master, text="Inclure les pré-inscrits", variable=self.Incl_Pre_inscrits)
        self.CheckPreInscrit.grid(row=2, column=0)
        
        self.CheckInvit = Checkbutton(master, text="Inclure les invités", variable=self.Incl_Invit)
        self.CheckInvit.grid(row=3, column=0)
        
        self.CheckInvit = Checkbutton(master, text="Ne pas limiter les doublons de mails", variable=self.KeepDuplicatesMails)
        self.CheckInvit.grid(row=4, column=0)
        
        self.close_button = Button(master, text="Quitter", command=master.quit)
        self.close_button.grid(row=4, column=2)

    def callback(self,url):
        #pour ouvrir l'aide
        webbrowser.open_new(url)

    def BBrowse(self):
        # Demande la selection du fichier à importer
        # initialise self.table = importation des données des contacts depuis l'intranet scout sous forme d'une liste 2D
        # info: l'intranet scout génère un fichier html renommé en .xls
        self.SourceFile.set(filedialog.askopenfilename())
        self.tableOK = True
        try:
            TableList = pd.read_html(self.SourceFile.get(),encoding = 'utf-8') # Returns list of all tables on page
            TableList = TableList[0].replace(np.nan, '', regex=True) # seule la table 0 existe. Je vire le Nan (= not a number)
            TableList =  TableList.sort_values(ColMaitrise, ascending=False) # on met en premier la maitrise
            self.table  = TableList.astype(str).values.tolist() # Table 0 W/O title
        except Exception as Err:
            self.tableOK = False
            print(str(Err))
        else:
            if len(self.table) < 1 or len(self.table[0]) < ColInscDateFin:
                self.tableOK = False
                self.InfoMsg("Le fichier source n'est pas bon:", "Error")
            if self.tableOK == True:
                del self.table[0] # efface la ligne de titre
                # création de la liste des unités du groupe
                self.TrouveUnites()
                self.InfoMsg("Fichier OK. Cliquer sur Go!", "Info")

    def TrouveUnites(self):
        # initialise le dictionnaire self.Structures avec les unités présentes dans le Groupe 
        # Format: {'1ERE LOUVETEAUX JEANNETTES - VAISE': 'Orange_1', 'AgeMin' : 7, 'Agemax':10}
        # analyse self.table pour lister tous les types d'unités du groupe
        ListStruct = []
        for i in range(len(self.table)):
            MaStruct = self.table[i][Colstruct]
            if MaStruct not in ListStruct:
                ListStruct.append(MaStruct)
        ListStruct.sort()
        # pour calculer l'année de naissance min/max , il faut savoir si on est avant ou après janvier
        Annee = datetime.datetime.now().year
        Annee = Annee  if datetime.datetime.now().month >= 9 else Annee -1
        # ListStruct contient la liste de tt les unités du groupe.
        #print(ListStruct)

        # initialisation des elements
        StrStructList='#'.join(ListStruct) # pour faciliter la recherche des unités en double ou triple
        for i in range(len(ListStruct)):
            StructFound = False
            for j in range(len(Struct_Type)):
                # cherche si on trouve la structure dans les structures connues
                if ListStruct[i].count(Struct_Type[j]['Txt']) >= 1 and not StructFound:
                    StructFound = True
                    MonLabel = Struct_Type[j]['Label']
                    AnneeMax = Annee - Struct_Type[j]['AgeMin']
                    AnneeMin =  Annee - Struct_Type[j]['AgeMax']
                    # si plus qu'une unité, on met le numéro de l'unité à la fin
                    if StrStructList.count(Struct_Type[j]['Txt']) >= 2:
                        MonLabel += "_" + ListStruct[i][:1]
                        #print(MonLabel)
            if not StructFound:
                MonLabel = ListStruct[i]
                AnneeMin = 1900
                AnneeMax = 2100
            self.Structures[ListStruct[i]] = [ MonLabel, AnneeMin, AnneeMax ]
       #print(self.Structures)
        
    def BGo(self):
        # création du dictionnaire de contacts + exports en fichier gmail/Excel
        if not self.tableOK:
            self.InfoMsg("Sélectionner d'abord un fichier valide", "Info")
        else:
            # Step 1: Lance la création du dictionnaire self.ContactsDic depuis self.table
            for i in range(len(self.table)): #range(55,65):  #                
                if (
                    (TxtPreInscrit in self.table[i][ColInscrit] and self.Incl_Pre_inscrits.get() == 1) or
                    (TxtInvit in self.table[i][ColInscrit] and self.Incl_Invit.get() == 1) or
                    TxtInscrit in self.table[i][ColInscrit]
                    ):
                    self.Add_Member(i)
                    if self.table[i][ColMaitrise] =='0':
                        self.Add_Parents(i)
            # tous les contacts sont ajoutés à self.ContactsDic
            # conversion en dataframe pandas pour pouvoir manipuler les lignes/colonnes plus facilement
            self.ContactsDF = pd.DataFrame.from_dict(self.ContactsDic)
            # prepare le DF
            # supprime colonne UID
            self.ContactsDF = self.ContactsDF.drop(columns=['UID'])
            # colonne Maitrise = non si vide ou '0' / Oui si '1'
            self.ContactsDF['Maitrise'] = ['Non' if Item != '1' else 'Oui' for Item in self.ContactsDF['Maitrise']]

            # lance l'export
            FNamePref = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S ")
            FName = FNamePref + '_GMAIL_ALL.csv'            
            self.ExportGmailCSV(FName,'All')

            for NomUnit in self.Structures:
                NomUnitCourt = self.Structures[NomUnit][0]
                FName = FNamePref + '_GMAIL_' + NomUnitCourt +'.csv'
                self.ExportGmailCSV(FName,NomUnitCourt)

            self.ExportExcel()
    
    def CreateNewMember(self):
        # création du template d'un dictionnaire d'un seul membre
        MemberDict = {}
        MemberDict['Nom'] = ''
        MemberDict['Prenom'] = ''
        MemberDict['Civ'] = ''
        for Key in self.Structures:
            MemberDict[self.Structures[Key][0]] = ''
        MemberDict['Maitrise'] = ''
        MemberDict['ParentType'] = ''
        MemberDict['Enfant'] = []
        MemberDict['InscType'] = ''
        MemberDict['InscDateFin'] = ''
        MemberDict['VerifAge'] = ''
        MemberDict['Fonction'] = ''
        MemberDict['FoncSecondaire'] = ''
        MemberDict['DateN'] = ''
        MemberDict['Mail1'] = ''
        MemberDict['Mail2'] = ''
        MemberDict['Tel1'] = ''
        MemberDict['Tel2'] = ''
        MemberDict['Tel3'] = ''
        MemberDict['Tel4'] = ''
        MemberDict['DroitImage'] = ''
        MemberDict['NumAlloc'] = ''
        MemberDict['NomLong'] = ''
        MemberDict['StructureLong'] = ''
        MemberDict['Adr'] = ''
        MemberDict['CP'] = ''
        MemberDict['Ville'] = ''
        MemberDict['Pays'] = ''
        MemberDict['UID'] = ''        
        return MemberDict.copy()

    def AddInContactDict(self, MemberDict):
        # ajoute le MemberDict au dictionnaire de contacts
        #print(MemberDict['Nom']+' '+MemberDict['Prenom'])
        if MemberDict['Nom'] != '':
            # Initialisation du dictionnaire
            if self.ContactsDic == {}:                 
                for key in MemberDict:
                    self.ContactsDic[key] = []
            # création d'un UID pour chercher les doublons de parents
            MemberDict['UID'] = MemberDict['Nom']+MemberDict['Prenom']+MemberDict['Tel1']+MemberDict['Tel2']+MemberDict['Tel3']+MemberDict['Tel4']
            if MemberDict['UID'] in self.ContactsDic['UID']:
                # l'ID existe déjà: c'est un parent (ou un parent de plusieurs enfant, ou un parent farfa ou un violet parent, ....).
                # on parcours en premier les membres donc on arrive ici uniquement pour les parents.
                # Il faut rajouter à la liste des enfants
                index = self.ContactsDic['UID'].index(MemberDict['UID'])
                self.ContactsDic['Enfant'][index] += MemberDict['Enfant']
                # il faut indiquer que le parent est parent:
                for Key in self.Structures:
                     # ajout des infos des colonnes structures
                     # ex: MemberDict['bleu']='Parent(s)'
                     NomUnitCourt=self.Structures[Key][0]
                     if MemberDict[NomUnitCourt] != '' and MemberDict[NomUnitCourt] not in self.ContactsDic[NomUnitCourt][index]:
                         if self.ContactsDic[NomUnitCourt][index] == '':
                             self.ContactsDic[NomUnitCourt][index] = MemberDict[NomUnitCourt]
                         else:    
                             self.ContactsDic[NomUnitCourt][index] += " et " + MemberDict[NomUnitCourt]
            else:
                # Nouveau contact
                for key in MemberDict:
                    #print(key)
                    self.ContactsDic[key].append(MemberDict[key])

        
    def Add_Member(self, i):
        # Ajoute le membre au Dictionnaire de Contacts
        # in: indic du tableau du fichier importé
        MemberDict = self.CreateNewMember()
        MemberDict['Civ'] = self.table[i][ColCiv]
        MemberDict['Nom'] = self.table[i][ColNom]
        MemberDict['Prenom'] = self.table[i][ColPrenom].title()
        MemberDict['NomLong'] = self.table[i][ColNom] + ' ' + self.table[i][ColPrenom].title()
        MemberDict['InscType'] = self.table[i][ColInscrit]
        MemberDict['Tel1'] = self.table[i][ColTel1]
        MemberDict['Tel2'] = self.table[i][ColTel2]
        MemberDict['Tel3'] = self.table[i][ColTel3]
        MemberDict['Tel4'] = self.table[i][ColTel4]
        MemberDict['DroitImage'] = self.table[i][ColDroitsImage]
        MemberDict['Adr'] = self.table[i][ColAdr1].title() + ' '+ self.table[i][ColAdr2].title() + ' ' + self.table[i][ColAdr3].title()
        MemberDict['CP'] = self.table[i][ColCP] 
        MemberDict['Ville'] = self.table[i][ColVille].title()
        MemberDict['Pays'] = self.table[i][ColPays].title()
        MemberDict['StructureLong'] = self.table[i][Colstruct]
        MaStruct = self.Structures[self.table[i][Colstruct]][0]
        if self.table[i][ColMaitrise] == '1':
            MemberDict[MaStruct] = LabelMaitrise
        else:
            MemberDict[MaStruct] = LabelMembre
        MemberDict['Maitrise'] = self.table[i][ColMaitrise]
        MemberDict['Fonction'] = self.table[i][ColFonction]
        
        if TxtCompaFnSecondr in self.table[i][ColFoncSecondaire] :
            # Compagnon fonction secondaire
            # on recherche la structure correspondante
            # ex:self.table[i][ColFoncSecondaire] = '140 - COMPAGNON (406933141 - 1ERE COMPAGNONS - VAISE)'
            TmpStr = self.table[i][ColFoncSecondaire]
            TmpStr=TmpStr[TmpStr.find("(")+1:TmpStr.find(")")] # supprime les parenthèses
            TmpStructCompa = TmpStr[TmpStr.find('-')+2:]
            MemberDict['FoncSecondaire'] = TmpStructCompa
            MemberDict[self.Structures[TmpStructCompa][0]]  = LabelMembre
        MemberDict['InscDateFin']= self.table[i][ColInscDateFin]
        MemberDict['DateN'] = self.table[i][ColDateN]
        MemberDict['NumAlloc'] = self.table[i][ColNumAlloc]
        # ajout des mails
        if self.KeepDuplicatesMails.get() == 1:
            MemberDict['Mail1'] = self.table[i][ColMail1]
            MemberDict['Mail2'] = self.table[i][ColMail2]            
        else:
            # on ne met pas le mail de l'enfant s'il est identique à celui des parents
            MailsParents = self.table[i][ColPapaMail1]+self.table[i][ColPapaMail2]+self.table[i][ColMamMail1]+self.table[i][ColMamMail2]
            if self.table[i][ColMail1] not in MailsParents:
                MemberDict['Mail1'] = self.table[i][ColMail1]
            if self.table[i][ColMail2] not in MailsParents:
                MemberDict['Mail2'] = self.table[i][ColMail2]
        
        # verifie l'age
        if self.table[i][ColMaitrise] != '1':
            AnneeNaiss = int(self.table[i][ColDateN][-4:])
            AnMax = self.Structures[self.table[i][Colstruct]][2]
            AnMin = self.Structures[self.table[i][Colstruct]][1]
            if AnneeNaiss > AnMax:
                MemberDict['VerifAge'] = 'Trop jeune?'
            elif AnneeNaiss < AnMin:
                MemberDict['VerifAge'] = 'Trop agé?'
            else:
                MemberDict['VerifAge'] = 'OK'
        self.AddInContactDict(MemberDict)

    def Add_Parents(self,i):
        # Ajoute les parents du membre au Dictionnaire de Contacts
        # in: indic du tableau du fichier importé
        MemberDict = self.CreateNewMember()
        MemberDict['Nom'] = self.table[i][ColPapaNom]
        MemberDict['Prenom'] = self.table[i][ColPapaPrenom].title()
        MemberDict['NomLong'] = self.table[i][ColPapaNom] + ' ' + self.table[i][ColPapaPrenom].title()
        MemberDict['Tel1'] = self.table[i][ColPapaTel1]
        MemberDict['Tel2'] = self.table[i][ColPapaTel2]
        MemberDict['Tel3'] = self.table[i][ColPapaTel3]
        MemberDict['Tel4'] = self.table[i][ColPapaTel4]
        MemberDict['Adr'] = self.table[i][ColPapaAdr1].title() + ' '+ self.table[i][ColPapaAdr2].title() + ' ' + self.table[i][ColPapaAdr3].title()
        MemberDict['CP'] = self.table[i][ColPapaCP] 
        MaStruct = self.Structures[self.table[i][Colstruct]][0] 
        MemberDict[MaStruct] = LabelParent
        MemberDict['Ville'] = self.table[i][ColPapaVille].title()
        MemberDict['Pays'] = self.table[i][ColPapaPays].title()
        # ajout des mails
        if self.KeepDuplicatesMails.get() == 1 :
            # on veut garder les mails des parents en double
            # soit par choix, soit parce que c'est un responsable ET parent (sinon on a une double entrée)
            MemberDict['Mail1'] = self.table[i][ColPapaMail1]
            MemberDict['Mail2'] = self.table[i][ColPapaMail2]            
        else:
            # on ne met pas le mail du papa s'il est identique à celui de maman
            MailsMaman = self.table[i][ColMamMail1]+self.table[i][ColMamMail2]
            if self.table[i][ColPapaMail1] not in MailsMaman:
                MemberDict['Mail1'] = self.table[i][ColPapaMail1]
            if self.table[i][ColPapaMail2] not in MailsMaman:
                MemberDict['Mail2'] = self.table[i][ColPapaMail2]
        MemberDict['ParentType'] = LabelPapa
        MemberDict['Civ'] = self.table[i][ColPapaCiv]
        if self.table[i][ColNom] != self.table[i][ColPapaNom]:
            Enfant = self.table[i][ColPrenom].title() +' '+ self.table[i][ColNom] + '[' + self.Structures[self.table[i][Colstruct]][0]+ ']'
        else:
            # même nom: on ne met que le prénom
            Enfant = self.table[i][ColPrenom].title() + '[' + self.Structures[self.table[i][Colstruct]][0]+ ']'
        MemberDict['Enfant'] = [ Enfant ]
        self.AddInContactDict(MemberDict.copy()) # ajoute Papa aux contacts
        
        MemberDict['Nom'] = self.table[i][ColMamNom]
        MemberDict['Prenom'] = self.table[i][ColMamPrenom].title()
        MemberDict['NomLong'] = self.table[i][ColMamNom] + ' ' + self.table[i][ColMamPrenom].title()
        MemberDict['Tel1'] = self.table[i][ColMamTel1]
        MemberDict['Tel2'] = self.table[i][ColMamTel2]
        MemberDict['Tel3'] = self.table[i][ColMamTel3]
        MemberDict['Tel4'] = self.table[i][ColMamTel4]
        MemberDict['Adr'] = self.table[i][ColMamAdr1].title() + ' '+ self.table[i][ColMamAdr2].title() + ' ' + self.table[i][ColMamAdr3].title()
        MemberDict['CP'] = self.table[i][ColMamCP] 
        MemberDict['Ville'] = self.table[i][ColMamVille].title()
        MemberDict['Pays'] = self.table[i][ColMamPays].title()
        MemberDict['Mail1'] = self.table[i][ColMamMail1]
        MemberDict['Mail2'] = self.table[i][ColMamMail2]
        MemberDict['ParentType'] = LabelMaman
        MemberDict['Civ'] = self.table[i][ColMamCiv]
        if self.table[i][ColNom] != self.table[i][ColMamNom]:
            Enfant = self.table[i][ColPrenom].title() +' '+ self.table[i][ColNom] + '[' + self.Structures[self.table[i][Colstruct]][0] + ']'
        else:
            Enfant = self.table[i][ColPrenom].title() + '[' + self.Structures[self.table[i][Colstruct]][0] + ']'
        MemberDict['Enfant'] = [ Enfant ]
        self.AddInContactDict(MemberDict.copy()) # ajoute Maman aux contacts


    def InfoMsg(self,Message,Type):
        # message d'info de la fenetre
        self.Info.set(Message)
        if Type == "Error":
            self.Infolabel.configure(bg="red")
        elif Type == "Success":  
            self.Infolabel.configure(bg="PaleGreen1")
        else:
            self.Infolabel.configure(bg="white")
        self.Infolabel.update()
        
    def PrepareCSV(self , MemberDict,FiltreUnit):
        # récupère un disctionnaire avec les infos à mettre
        # retourne une liste avec la ou les lignes pour export en CSV sous forme de liste
        CSV_Row_List = [''] * CSVNbCol # crée une liste vide par défaut
        OutList = []
        # liste où ajouter
        CSV_Row_List[CSV_Prenom] = MemberDict['Prenom']
        CSV_Row_List[CSV_Nom] = MemberDict['Nom']
        if len(MemberDict['Enfant']) >= 1:
            Suffixe = ' (' + ' '.join(MemberDict['Enfant']) +')'
        else :
            Suffixe = ''
        CSV_Row_List[CSV_notes] = MemberDict['InscType'] +' jusqu\'au ' + MemberDict['InscDateFin']
        CSV_Row_List[CSV_tel1] = MemberDict['Tel1']
        CSV_Row_List[CSV_tel2] = MemberDict['Tel2']
        CSV_Row_List[CSV_tel3] = MemberDict['Tel3']
        CSV_Row_List[CSV_tel4] = MemberDict['Tel4']
        CSV_Row_List[CSV_TelType1] = 'Tel1'
        CSV_Row_List[CSV_TelType2] = 'Tel2'
        CSV_Row_List[CSV_TelType3] = 'Tel3'
        CSV_Row_List[CSV_TelType4] = 'Tel4'
        CSV_Row_List[CSV_Adr1] = MemberDict['Adr']
        CSV_Row_List[CSV_ville] = MemberDict['Ville']
        CSV_Row_List[CSV_CP] = MemberDict['CP']
        CSV_Row_List[CSV_Pays] = MemberDict['Pays']
#         CSV_Row_List[CSV_Structure] = MemberDict['Structure']
        
        # maj des étiquettes GMAIL
        CSV_Row_List[CSV_Group] = '* My Contacts'
        for NomUnit in self.Structures:
            NomUnitCourt = self.Structures[NomUnit][0]
            TypeMembre = MemberDict[NomUnitCourt] # type membre = LabelMaitrise, ou LabelParent, ...
            if TypeMembre != '':
                if FiltreUnit == 'All' or NomUnitCourt == FiltreUnit:
                    # attention, type membre peut être "Maitrise Parent(s)"
                    if LabelParent in TypeMembre:
                        CSV_Row_List[CSV_Group] += ' ::: ' + NomUnitCourt + ' ' + LabelParent
                    if LabelMaitrise in TypeMembre:
                        CSV_Row_List[CSV_Group] += ' ::: ' + NomUnitCourt + ' ' + LabelMaitrise
                    if LabelMembre in TypeMembre:
                        CSV_Row_List[CSV_Group] += ' ::: ' + NomUnitCourt + ' ' + LabelMembre
                # ajout des étiquettes globales
                if FiltreUnit == 'All':
                    if (
                        (LabelMembre in CSV_Row_List[CSV_Group] or LabelMaitrise in CSV_Row_List[CSV_Group])
                        and LabelTouslesMembres not in CSV_Row_List[CSV_Group]):
                            CSV_Row_List[CSV_Group] += ' ::: ' + LabelTouslesMembres
                    if LabelParent in CSV_Row_List[CSV_Group] and LabelTousLesParents not in CSV_Row_List[CSV_Group]:
                        CSV_Row_List[CSV_Group] += ' ::: ' + LabelTousLesParents
#                     if LabelMaitrise in CSV_Row_List[CSV_Group] and SelectLabelChefs in NomUnitCourt and LabelChefs not in CSV_Row_List[CSV_Group]:
#                         CSV_Row_List[CSV_Group] += ' ::: ' + LabelChefs
                    # print(NomUnitCourt)
                    #    print(SelectLabelChefs) # =OrangesBleusRouges

        # on crée un contact par email = 2 contacts si on a 2 mails
        Mail1_OK = self.CheckMail(MemberDict['Mail1'])
        Mail2_OK = self.CheckMail(MemberDict['Mail2'])
        if Mail1_OK and Mail2_OK:
            CSV_Row_List[CSV_mail] = MemberDict['Mail1']
            CSV_Row_List[CSV_Suffixe] = Suffixe + ' Mail1'
            CSV_Row_List[CSV_NomLong] = MemberDict['NomLong'] + CSV_Row_List[CSV_Suffixe]
            OutList.append(CSV_Row_List.copy()) # le .copy permet de copier une liste indépendante. Sinon, on copie un pointeur --> doublons

            CSV_Row_List[CSV_mail] = MemberDict['Mail2']
            CSV_Row_List[CSV_Suffixe] = Suffixe + ' Mail2'
        else:
            CSV_Row_List[CSV_Suffixe] = Suffixe
            if Mail1_OK:
                CSV_Row_List[CSV_mail] = MemberDict['Mail1']
            else:
                CSV_Row_List[CSV_mail] = MemberDict['Mail2']
        CSV_Row_List[CSV_NomLong] = MemberDict['NomLong'] + CSV_Row_List[CSV_Suffixe]
        OutList.append(CSV_Row_List.copy()) # le .copy permet de copier une liste indépendante. Sinon, on copie un pointeur --> doublons
        return OutList
      
    def ExportGmailCSV(self, FName, FiltreUnit):
        # Création du CSV pour Gmail
        # FiltreUnit = All pour tout intégrer, sinon le nom court de l'unité.
        # Step 1: on crée une liste 2D avec tous les champs nécéssaire à partir du dataframe
        ContactsCSV  = [] # liste 2D à exporter en CSV au format gmail. entete dans la "CSV_Header"
        if FiltreUnit == 'All':
            DF = self.ContactsDF
        else:
            DF = self.ContactsDF[(self.ContactsDF[FiltreUnit] != '')  ]
        for i in range(DF.shape[0]):
            Contact = DF.iloc[i]
            ContactsCSV += self.PrepareCSV(Contact,FiltreUnit).copy()
        # Step 2: création du fichier
        MyDir = os.path.dirname(self.SourceFile.get())
        CSV_Name = MyDir+"/"+ FName 
        try:
            with open(CSV_Name, "w", newline="", encoding="utf-8") as f:
                f.write(CSV_Header + '\n')
                writer = csv.writer(f, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                writer.writerows(ContactsCSV)
                f.close()
        except:
            self.InfoMsg("Erreur dans l'export du fichier CSV" + MyDir, "Error")
        else:
            self.InfoMsg("Fichier contact généré ici: " + MyDir, "Success")
        
    def ExportExcel(self):
        MyDir = os.path.dirname(self.SourceFile.get())
        XlsName = MyDir+"/"+ datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S ") 
        XlsNameGlobal = XlsName + 'SGDF.xlsx'
        ew = StyleFrame.ExcelWriter(XlsNameGlobal)
        #exporte la liste de tous les contacts du groupe
        df1 = self.ContactsDF.drop(columns=['NomLong']).copy()
        # améliore la présentation des enfants
        df1['Enfant'] = [ ' \n'.join(Item) for Item in self.ContactsDF['Enfant']]
        df1 = df1.sort_values(by=['Nom', 'Prenom'])
        ListColumns = df1.columns.values.tolist()
        sf1 = StyleFrame(df1)
        sf1.to_excel(excel_writer=ew, 
                    row_to_add_filters=0, 
                    best_fit=ListColumns, 
                    columns_and_rows_to_freeze='C2',
                    sheet_name="Tous")
        TableauBord = []
        TableauBord_Col = ['Structure', 'Incrits', 'invités', 'Filles', 'Garçons', 'Total' ]
        for Structure in self.Structures:
            # Listing des membres de chaque unité
            StructLabel = self.Structures[Structure][0]
            #StructDF = df1[df1['StructureLong'].str.contains(Structure) | df1['FoncSecondaire'].str.contains(Structure)].copy()
            StructDF = df1[df1[StructLabel] != '' ].copy()
            StructDF = StructDF[(StructDF[StructLabel] != LabelParent)  ] # selectionne tt le monde sauf les parents
            StructDF = StructDF[['Civ', 'Nom','Prenom','Maitrise','InscType', 'VerifAge', 'NumAlloc',  'DroitImage', 'DateN']]
            StructDF = StructDF.sort_values(by=['Maitrise','Nom', 'Prenom'])
            # print(StructDF)
            if not StructDF.empty:
                # pas d'export s'il n'y a personne dans l'unité
                ListColumns = StructDF.columns.values.tolist()
                sf2 = StyleFrame(StructDF)
                sf2.to_excel(excel_writer=ew,
                             row_to_add_filters=0,
                             best_fit=ListColumns,
                             columns_and_rows_to_freeze='C2',
                             sheet_name=StructLabel)
                # ecriture de l'excel de chaque unité
                ew2 = StyleFrame.ExcelWriter(XlsName + ' ' + StructLabel + '.xlsx')
                sf2.to_excel(excel_writer=ew2,
                             row_to_add_filters=0,
                             best_fit=ListColumns,
                             columns_and_rows_to_freeze='C2',
                             sheet_name=StructLabel)
                ew2.close()
            
            # Compte le nombre d'enfants:
            StructDF = StructDF[(StructDF['Maitrise'] == 'Non' ) ]
            StructDF = StructDF[(~StructDF['InscType'].str.contains(TxtPreInscrit) ) ] # ne contient pas PréInscrit
            NbInscrit = StructDF[StructDF['InscType'].str.contains(TxtInscrit)].shape[0]
            NbInvit = StructDF[StructDF['InscType'].str.contains(TxtInvit)].shape[0]
            NbFille = StructDF[StructDF['Civ'].str.contains(TxtFille)].shape[0]
            NbTot = NbInvit + NbInscrit #+ NbPreInscr
            NbGar = NbTot - NbFille
            if NbTot != 0:
                TableauBord.append([Structure, NbInscrit, NbInvit, NbFille, NbGar, NbInscrit+NbInvit])
        # ajoute le tableau de bord à l'excel
        TableauBordDF = pd.DataFrame(TableauBord, columns=TableauBord_Col)
        TableauBordSF = StyleFrame(TableauBordDF)
        ListColumns = TableauBordDF.columns.values.tolist()
        TableauBordSF.to_excel(excel_writer=ew, 
                     row_to_add_filters=0, 
                     best_fit=ListColumns, 
                     sheet_name="TableauBord")
        ew.close()

    def CheckMail(self,Mail):
        # vérifie si mail valide
        regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'     
        if re.fullmatch(regex, Mail):
            return True
        else:
            return False

                
root = Tk()
my_gui = SGDF_ToolboxUI(root)
root.mainloop()
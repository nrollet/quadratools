import pyodbc
import logging
import sys
import random
import string
import os
import shutil
from datetime import datetime
from quadratools.sqlstr import req_calc_central

def progressbar(count, total):
    """
    Pour l'affichage d'une barre de progression
    pendant l'insert des écritures
    """
    bar = "#"*40
    level = int((count*40)/total)+1
    tail = "\r"
    if count == total:
        tail = "\n"

    print("[{}] {}/{}".format(bar[0:level].ljust(40),
                            str(count).zfill(len(str(total))),
                            str(total)), end=tail)

def gen_image_name(reserved):
    """
    Création aléatoire de nom pour les images pieces
    Avec vérif si le fichier existe déjà
    """
    name = "".join(random.choice(string.ascii_uppercase) for _ in range(10))
    if name in reserved:
        gen_image_name(reserved)
    else:
        return name

def doc_rename(filename):
    """
    Outil pour renommer un document avec un nom aléatoire
    """
    salt = "".join(random.choice(string.ascii_lowercase) for _ in range(3))
    splitbase = filename.split(".")
    if len(splitbase) > 1:
        base = splitbase[0] + "_" + salt + "." + splitbase[1]
    else:
        base = splitbase[0] + "_" + salt
    return base

class QueryCompta(object):

    def __init__(self):

        self.rs = ""
        self.exedeb = datetime(year=1800, month=1, day=1)
        self.exefin = datetime(year=1800, month=1, day=1)
        self.dtsortie = datetime(year=1800, month=1, day=1)
        self.dtvalid = datetime(year=1800, month=1, day=1)
        self.dtclot = datetime(year=1800, month=1, day=1)
        self.collfrn = "40100000"
        self.collcli = "41100000"
        self.preffrn = "0"
        self.prefcli = "9"
        self.plan = {}
        self.journaux = []
        self.affect = {}
        self.images = []

    def connect(self, chem_base):
        self.chem_base = chem_base.lower()

        constr = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=' + \
            self.chem_base
        try:
            self.conx = pyodbc.connect(constr, autocommit=True)
            logging.info('Ouverture de {}'.format(self.chem_base))
            self.cursor = self.conx.cursor()

        except pyodbc.Error:
            logging.error("erreur requete base {} \n {}".format(
                self.chem_base, sys.exc_info()[1]))
        except:
            logging.error("erreur ouverture base {} \n {}".format(
                self.chem_base, sys.exc_info()[0]))

        # Paramètres table Dossier1
        sql = """
            SELECT
            RaisonSociale, DebutExercice, FinExercice,
            PeriodeValidee, PeriodeCloturee
            FROM Dossier1
        """
        self.cursor.execute(sql)
        for rs, exd, exf, pv, pc in self.cursor.fetchall():
            self.rs = rs
            self.exedeb = exd
            self.exefin = exf
            self.dtvalid = pv
            self.dtclot = pc

        # Paramètres table Dossier2
        sql = """
            SELECT
            CollectifFrnDefaut, CollectifClientDefaut,
            CodifClasse0Seule, CodifClasse9Seule, DSDateSortie
            FROM Dossier2
        """
        self.cursor.execute(sql)
        for colfrn, colcli, cl0, _, datesortie in self.cursor.fetchall():
            self.collfrn = colfrn
            self.collcli = colcli
            self.dtsortie = datesortie
            if cl0 == "C":
                self.preffrn = "9"
                self.prefcli = "0"

        # Plan comptable
        sql = """
            SELECT
            Numero, Type, Intitule, NbEcritures, ProchaineLettre
            FROM Comptes
        """
        self.cursor.execute(sql)
        for num, _, intit, nbecr, lettr in self.cursor.fetchall():
            self.plan.update(
                {
                    num : {
                        "intitule" : intit,
                        "nbecr" : nbecr,
                        "lettrage" : lettr
                    }
            })

        # Codes journaux
        sql = """SELECT Code from Journaux"""
        self.cursor.execute(sql)
        for cj in self.cursor.fetchall():
            self.journaux.append(cj)        

        # Affectations Ana
        sql = """SELECT NumCompte, CodeCentre FROM AffectationAna"""
        self.cursor.execute(sql)
        for compte, centre in self.cursor.fetchall():
            self.affect.update(
                {compte : centre}
            )

        # Liste des images existantes
        sql = """SELECT DISTINCT RefImage FROM Ecritures"""
        self.cursor.execute(sql)
        for image in self.cursor.fetchall():
            self.images.append(image[0].split(".")[0])

        logging.info("Nom : {}".format(self.rs))
        logging.info("Exercice : {} {}".format(
            self.exedeb.strftime("%d/%m/%Y"),
            self.exefin.strftime("%d/%m/%Y")))
        logging.info("Coll. fourn : {}".format(self.collfrn))
        logging.info("Préfixe fourn : {}".format(self.preffrn))
        logging.info("Préfixe client : {}".format(self.prefcli))
        logging.info("Comptes : {}".format(len(self.plan)))
        logging.info("Période cloturée : {}".format(self.dtclot.strftime("%d/%m/%Y")))  
        logging.info("Date sortie : {}".format(self.dtsortie.strftime("%d/%m/%Y")))  

    def close(self):
        logging.info('fermeture de la base')
        self.conx.commit()
        self.conx.close()

    def exec_select(self, sql_string):
        """
        Pour exécuter une requete sql passée en argument
        """
        data = ""
        try:
            data = self.cursor.execute(sql_string).fetchall()
        except pyodbc.Error:
            logging.error("erreur requete base {} \n {}".format(
                self.chem_base, sys.exc_info()[1]))
            logging.error("Requete SQL : {}".format(sql_string))                
        except:
            logging.error("erreur ouverture base {} \n {}".format(
                self.chem_base, sys.exc_info()[0]))
            logging.error("Requete SQL : {}".format(sql_string))
        return data

    def exec_insert(self, sql_string):
        """
        Pour exécuter une requete sql passée en argument
        """
        status = False
        try:
            self.cursor.execute(sql_string)
            status = True
        except pyodbc.Error:
            logging.error("erreur requete base {} \n {}".format(
                self.chem_base, sys.exc_info()[1]))
            logging.error("Requete SQL : {}".format(sql_string))                
        except:
            logging.error("erreur ouverture base {} \n {}".format(
                self.chem_base, sys.exc_info()[0]))
            logging.error("Requete SQL : {}".format(sql_string))
        return status


    def journal(self, code_journal):
        """
        Récupère le solde de tous les comptes d'un journal donné
        renvoi dictionnaire dont la clé est le n° de pièce
        """

        sql = """
            SELECT
            E.NumeroPiece, E.NumeroCompte, C.Intitule, E.PeriodeEcriture,
            SUM(E.MontantTenuDebit-E.MontantTenuCredit) AS Solde
            FROM Ecritures E
            LEFT JOIN Comptes C ON E.NumeroCompte=C.Numero
            WHERE CodeJournal='{}'
            AND TypeLigne='E'
            GROUP BY NumeroPiece, NumeroCompte, Intitule, PeriodeEcriture
        """.format(code_journal)

        data = self.cursor.execute(sql).fetchall()
        dic = {}
        for piece, compte, intitule, periode, solde in data:
            dic.setdefault(compte, {})
            dic[compte].update({"intitule": intitule})
            dic[compte].setdefault("piece", [])
            dic[compte]["piece"].append([piece, periode, solde])
        return dic

    def get_solde_compte(self, compte):

        sql = """
            SELECT
            SUM(MontantTenuDebit) AS Debit,
            SUM(MontantTenuCredit) AS Credit
            FROM Ecritures
            WHERE NumeroCompte='{}'
            AND TypeLigne='E'
        """.format(compte)
        self.cursor.execute(sql)
        return self.cursor.fetchall()[0]

    def insert_compte(self, compte):
        """
        Ajout d'un nouveau compte dans la table Comptes
        """

        # Détection de la nature du compte
        if compte.startswith(self.preffrn):
            type_cpt = "F"
            Collectif = self.collfrn
        elif compte.startswith(self.prefcli):
            type_cpt = "C"
            Collectif = self.collcli
        elif compte[0] in ["1", "2", "3", "4", "5", "6", "7"]:
            type_cpt = "G"
            Collectif = "NULL"
        else:
            logging.error("Compte {} hors PC".format(compte))
            return False

        Collectif = "'{}'".format(Collectif)

        TypeCollectif = False
        Debit = 0.0
        Credit = 0.0
        DebitHorsEx = 0.0
        CreditHorsEx = 0.0
        Debit_1 = 0.0
        Credit_1 = 0.0
        Collectif_1 = Collectif
        Debit_2 = 0.0
        Credit_2 = 0.0
        Collectif_2 = Collectif
        NbEcritures = False
        DetailCloture = 1
        ALettrerAuto = 1
        CentraliseGdLivre = False
        SuiviQuantite = False
        CumulPiedJournal = False
        TvaEncaissement = False
        DateSysCreation = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ProchaineLettre = "AA"
        NoProchainLettrage = False
        CodeTva = False
        EditM2 = False
        Franchise = False
        IntraCom = False
        NiveauDroit = 0
        GererIntCptCour = False
        MargeTheorique = 100
        TvaDOM = False
        Periodicite = False
        SuiviDevises = False
        CompteInactif = False
        BonAPayer = False
        QuantiteNbEntier = 5
        QuantiteNbDec = 2
        PrixMoyenNbEntier = 5
        PrixMoyenNbDec = 2
        PersonneMorale = False
        SuiviQuantite2 = False
        QuantiteNbEntier2 = 5
        QuantiteNbDec2 = 2
        PrixMoyenNbEntier2 = 5
        PrixMoyenNbDec2 = 2
        RepartitionAna = 0
        RepartitionAuto = False
        ActiverLotTrace = False
        TvaAutresOpeImpos = False
        ProchaineLettreTiers = "AAA"
        PrestaTel = False
        TypeIntraCom = 0
        Prestataire = 0
        CptParticulier = False

        logging.info("Ajout d'un nouveau compte : {} ({})".format(compte, type_cpt))

        sql = f"""
            INSERT INTO Comptes
            (Numero, Type, TypeCollectif,
            Debit, Credit, DebitHorsEx, CreditHorsEx,
            Collectif,
            Debit_1, Credit_1, Collectif_1,
            Debit_2, Credit_2, Collectif_2,
            NbEcritures, DetailCloture,
            ALettrerAuto, CentraliseGdLivre,
            SuiviQuantite, CumulPiedJournal,
            TvaEncaissement, DateSysCreation,
            ProchaineLettre, NoProchainLettrage,
            CodeTva,EditM2, Franchise,
            IntraCom, NiveauDroit, GererIntCptCour,
            MargeTheorique, TvaDOM, Periodicite,
            SuiviDevises, CompteInactif, BonAPayer,
            QuantiteNbEntier, QuantiteNbDec,
            PrixMoyenNbEntier, PrixMoyenNbDec,
            PersonneMorale, SuiviQuantite2,
            QuantiteNbEntier2, QuantiteNbDec2,
            PrixMoyenNbEntier2, PrixMoyenNbDec2,
            RepartitionAna, RepartitionAuto,
            ActiverLotTrace, TvaAutresOpeImpos,
            ProchaineLettreTiers, PrestaTel,
            TypeIntraCom,Prestataire,CptParticulier)
            VALUES
            ('{compte}', '{type_cpt}', {TypeCollectif},
            {Debit}, {Credit}, {DebitHorsEx}, {CreditHorsEx},
            {Collectif}, 
            {Debit_1}, {Credit_1}, {Collectif_1},
            {Debit_2}, {Credit_2}, {Collectif_2},
            {NbEcritures}, {DetailCloture},
            {ALettrerAuto}, {CentraliseGdLivre},
            {SuiviQuantite}, {CumulPiedJournal},
            {TvaEncaissement}, #{DateSysCreation}#,
            '{ProchaineLettre}', {NoProchainLettrage},
            {CodeTva}, {EditM2}, {Franchise},
            {IntraCom}, {NiveauDroit}, {GererIntCptCour},
            {MargeTheorique}, {TvaDOM}, {Periodicite},
            {SuiviDevises}, {CompteInactif}, {BonAPayer},
            {QuantiteNbEntier}, {QuantiteNbDec},
            {PrixMoyenNbEntier}, {PrixMoyenNbDec},
            {PersonneMorale}, {SuiviQuantite2},
            {QuantiteNbEntier2}, {QuantiteNbDec2},
            {PrixMoyenNbEntier2}, {PrixMoyenNbDec2},
            {RepartitionAna}, {RepartitionAuto},
            {ActiverLotTrace}, {TvaAutresOpeImpos},
            '{ProchaineLettreTiers}', {PrestaTel},
            {TypeIntraCom}, {Prestataire}, {CptParticulier})
        """
        try:
            self.cursor.execute(sql)
            logging.info("Mise à jour table Comptes OK")
            self.plan.update({compte : ""})
            return True
        except pyodbc.Error:
            logging.debug(sql)
            logging.error("erreur insert Comptes {}".format(sys.exc_info()[1]))
            return False

    def get_last_uniq(self):
        """
        Retourne le dernier identifiant unique
        de la liste des écritures
        """
        sSQL = "SELECT MAX(NumUniq) FROM Ecritures"
        self.cursor.execute(sSQL)
        numero = self.cursor.fetchall()[0][0]
        if numero == None: numero = 0
        return numero

    def get_last_lignefolio(self, journal, periode):
        """
        Retourne le dernier numero de ligne folio
        utilisé dans un journal
        """
        sSQL = ("SELECT MAX(LigneFolio) FROM Ecritures "
                "WHERE CodeJournal='{}' "
                "AND PeriodeEcriture=#{}#").format(journal, periode)
        self.cursor.execute(sSQL)
        self.cursor.execute(sSQL)
        last_lfolio = list(self.cursor)[0][0]  
        if last_lfolio is None: last_lfolio = 0
        return last_lfolio             

    def insert_ecrit(self, compte, journal, folio, date,
                     libelle, debit, credit, piece, image, centre, image_root):
        """
        Insere une nouvelle ligne dans la table ecritures de Quadra.
        Si le compte possède une affectation analytique, une deuxème
        ligne est insérée avec les données analytiques
        """
        # Contrôle piece-jointe
        source_image_path = os.path.join(image_root, image)
        if image and not os.path.isfile(source_image_path):
            logging.error("fichier absent {}".format(os.path.join(image_root, image)))
            image = ""

        # Folio = "000"
        MontantSaisiDebit = 0
        MontantSaisiCredit = 0
        Quantite = 0
        NumLettrage = 0
        RapproBancaireOk = False
        NoLotEcritures = 0
        PieceInterne = 0
        CodeOperateur = "QTOO"
        Etat = 0
        NumLigne = 0
        TypeLigne = "E"
        Actif = False
        PrctRepartition = 0
        ClientOuFrn = 0
        MontantAna = 0
        MilliemesAna = 0
        CodeTva = -1
        BonsAPayer = False
        MtDevForce = False
        EnLitige = False
        Quantite2 = 0
        NumEcrEco = 0
        NoLotFactor = 0
        Validee = False
        NoLotIs = 0
        NumMandat = 0
        DateSysSaisie = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        CompteContrepartie = ""   

        # Elements analytique invar.
        A_NumLigne = 1
        A_TypeLigne = "A"
        A_Nature = '*'
        A_PrctRepartition = 100
        A_TypeSaisie = 'P'  

        periode = datetime(date.year,
                           date.month,1)
        jour = date.day
        libelle = libelle[:30]
        piece = piece[0:10]

        uid = self.get_last_uniq() + 1
        lfolio = self.get_last_lignefolio(journal, periode) + 10

        # Vérif date sur période non cloturée
        if periode <= self.dtclot:
            logging.error("ecriture sur periode cloturée : {}".format(periode))
            return False

        # Vérif affect analytique
        if not centre:
            if compte in self.affect.keys():
                centre = self.affect[compte]

        # Vérif existence du compte
        if not compte in self.plan:
            # Si insert ok on met à jour les param_doss
            if self.insert_compte(compte):
                self.plan.update({compte : {"initule": "", "nbrecr": 0, "lettrage": ""}})

        sql = f"""
            INSERT INTO Ecritures
            (NumUniq, NumeroCompte, 
            CodeJournal, Folio, 
            LigneFolio, PeriodeEcriture, 
            JourEcriture, Libelle, 
            MontantTenuDebit, MontantTenuCredit, 
            MontantSaisiDebit, MontantSaisiCredit, 
            CompteContrepartie, Quantite, 
            NumeroPiece, NumLettrage, 
            RapproBancaireOk, NoLotEcritures, 
            PieceInterne, CodeOperateur, 
            DateSysSaisie, Etat, 
            NumLigne, TypeLigne, 
            Actif, PrctRepartition, 
            ClientOuFrn, RefImage, 
            MontantAna, MilliemesAna, 
            CentreSimple, CodeTva, 
            BonsAPayer, MtDevForce, 
            EnLitige, Quantite2, 
            NumEcrEco, NoLotFactor, 
            Validee, NoLotIs, 
            NumMandat) 
            VALUES 
            ({uid}, '{compte}', 
            '{journal}', {folio}, 
            {lfolio}, #{periode}#, 
            {jour}, '{libelle}', 
            {debit}, {credit}, 
            {MontantSaisiDebit}, {MontantSaisiCredit}, 
            '{CompteContrepartie}', {Quantite}, 
            '{piece}', {NumLettrage}, 
            {RapproBancaireOk}, {NoLotEcritures}, 
            {PieceInterne}, '{CodeOperateur}', 
            #{DateSysSaisie}#, {Etat}, 
            {NumLigne}, '{TypeLigne}', 
            {Actif}, {PrctRepartition}, 
            {ClientOuFrn}, '{image}', 
            {MontantAna}, {MilliemesAna}, 
            '{centre}', {CodeTva}, 
            {BonsAPayer}, {MtDevForce}, 
            {EnLitige}, {Quantite2}, 
            {NumEcrEco}, {NoLotFactor}, 
            {Validee}, {NoLotIs}, 
            {NumMandat})
            """
        try:
            self.cursor.execute(sql)
        except pyodbc.Error:
            logging.error("erreur insert Ecritures {}".format(sys.exc_info()[1]))
            logging.debug(sql)
            return False

        if centre:
            montant = abs(debit - credit)
            uid += 1

            sql = f"""
                INSERT INTO Ecritures
                (NumUniq, NumeroCompte,
                CodeJournal, Folio,
                LigneFolio,PeriodeEcriture,
                JourEcriture, MontantTenuDebit,
                MontantTenuCredit, MontantSaisiDebit,
                MontantSaisiCredit, Quantite,
                NumLettrage, RapproBancaireOk,
                NoLotEcritures, PieceInterne,
                NumLigne, TypeLigne,
                Actif, Centre,
                Nature, PrctRepartition,
                TypeSaisie, MontantAna,
                MilliemesAna, CodeTva,
                BonsAPayer, MtDevForce,
                EnLitige, Quantite2,
                NumEcrEco, NoLotFactor,
                Validee, NoLotIs,
                NumMandat)
                VALUES
                ({uid}, '{compte}',
                '{journal}', {folio},
                {lfolio}, #{periode}#,
                {jour}, {debit}, 
                {credit}, {MontantSaisiDebit},
                {MontantSaisiCredit}, {Quantite},
                {NumLettrage}, {RapproBancaireOk},
                {NoLotEcritures}, {PieceInterne},
                {A_NumLigne}, '{A_TypeLigne}',
                {Actif}, '{centre}',
                '{A_Nature}', {A_PrctRepartition},
                '{A_TypeSaisie}', {montant},
                {MilliemesAna}, {CodeTva},
                {BonsAPayer}, {MtDevForce},
                {EnLitige}, {Quantite2},
                {NumEcrEco}, {NoLotFactor},
                {Validee}, {NoLotIs}, {NumMandat})
                """           
            try:
                self.cursor.execute(sql)
            except pyodbc.Error:
                logging.error("erreur insert Ecritures {} (A)".format(sys.exc_info()[1]))
                logging.debug(sql)
                return False

            # TRAITEMENT IMAGE
            dest_image_dir = self.chem_base.replace("qcompta.mdb", "images")
            if os.path.isfile(source_image_path):
                if os.path.isdir(dest_image_dir): 
                    try:  
                        shutil.copy(source_image_path, dest_image_dir+"/"+image)  
                    except IOError as e:
                        logging.error("Echec copie {}, {}".format(source_image_path, e))

            return uid
    
    def calc_centralisateurs(self):
        sql = req_calc_central(self.prefcli, self.preffrn)
        data = self.exec_select(sql)

        # modification sur la requête :
        for index, row in enumerate(data):
            # La valeurs None -> 0
            row = [0.0 if x==None else x for x in row]
            # Mystérieusement la date repasse en string... Il faut la remettre en datetime
            # if row[1]:
            #     row[1] = datetime.strptime(row[1], "%d/%m/%Y")
            # Calcul du numéro la prochaine ligne dans le folio (à la dizaine supérieure)
            if row[4]:
                row[4] = (int(row[4]/10)+1)*10
            data[index] = row
        return data      

    def maj_centralisateurs(self):
        """
        Mise à jour de toute la table des centralisateurs
        """
        logging.info("purge des centralisateurs")
        self.exec_select("DELETE FROM Centralisateur WHERE 1")
        data = self.calc_centralisateurs()

        logging.info("Mise à jour de la table Centralisateurs")
        count = 1

        for (journal, periode, folio,
             nbligne, prligne, 
             debitcli, creditcli,
             debitfrn, creditfrn, 
             debit15, credit15, 
             debit67, credit67) in data:

            logging.debug("insert central {} {}".format(journal, periode))
            sql = f"""
                INSERT INTO Centralisateur 
                (CodeJournal, Periode, Folio, 
                NbLigneFolio, ProchaineLigne, 
                DebitClient, CreditClient, 
                DebitFournisseur, CreditFournisseur, 
                DebitClasse15, CreditClasse15, 
                DebitClasse67, CreditClasse67) 
                VALUES 
                ('{journal}', #{periode}#, {folio}, 
                {nbligne}, {prligne}, 
                {debitcli}, {creditcli}, 
                {debitfrn}, {creditfrn}, 
                {debit15}, {credit15}, 
                {debit67}, {credit67})
            """
            self.exec_insert(sql)
            
        logging.info("Mise à jour table Centralisateurs OK")
        return True                        

    def calc_solde_comptes(self):
        """
        Requête pour calculer les soldes de tous les comptes
        a partir de la table écriture
        """
        sql = f"""
        SELECT N.NumeroCompte,
            N.debit, N.credit,
            N1.debit, N1.credit,
            N.NbEcritures
        FROM (
            SELECT NumeroCompte, 
            ROUND(SUM(MontantTenuDebit),2) AS debit, 
            ROUND(SUM(MontantTenuCredit),2) AS credit, 
            COUNT(*) AS NbEcritures
            FROM Ecritures
            WHERE TypeLigne='E'
            GROUP BY NumeroCompte) N            
            LEFT JOIN (
            SELECT NumeroCompte, 
            ROUND(SUM(MontantTenuDebit),2) AS debit, 
            ROUND(SUM(MontantTenuCredit),2) AS credit
            FROM Ecritures
            WHERE PeriodeEcriture>=#{self.exefin}# 
            AND TypeLigne='E'
            GROUP BY NumeroCompte) N1
            ON N.NumeroCompte=N1.NumeroCompte
        UNION
        SELECT N.Compte,
            N.debit, N.credit,
            N1.debit, N1.credit,
            N.NbEcritures
        FROM ((
            SELECT
            '{self.collfrn}' AS Compte, COUNT(*) AS NbEcritures,
            ROUND(SUM(MontantTenuDebit),2) AS debit, 
            ROUND(SUM(MontantTenuCredit),2) as credit
            FROM Ecritures
            WHERE TypeLigne='E'
            AND NumeroCompte LIKE '0%') N
            LEFT JOIN (
            SELECT
            '{self.collfrn}' AS Compte, COUNT(*) AS NbEcritures,
            ROUND(SUM(MontantTenuDebit),2) AS debit, 
            ROUND(SUM(MontantTenuCredit),2) AS credit
            FROM Ecritures
            WHERE TypeLigne='E'
            AND NumeroCompte LIKE '0%'
            AND PeriodeEcriture>=#{self.exefin}#) N1
            ON N.Compte=N1.Compte)
        UNION
        SELECT N.Compte,
            N.debit, N.credit,
            N1.debit, N1.credit,
            N.NbEcritures
        FROM ((
            SELECT
            '{self.collcli}' AS Compte, COUNT(*) AS NbEcritures,
            ROUND(SUM(MontantTenuDebit),2) AS debit, 
            ROUND(SUM(MontantTenuCredit),2) as credit
            FROM Ecritures
            WHERE TypeLigne='E'
            AND NumeroCompte LIKE '9%') N
            LEFT JOIN (
            SELECT
            '{self.collcli}' AS Compte, COUNT(*) AS NbEcritures,
            ROUND(SUM(MontantTenuDebit),2) AS debit, 
            ROUND(SUM(MontantTenuCredit),2) AS credit
            FROM Ecritures
            WHERE TypeLigne='E'
            AND NumeroCompte LIKE '9%'
            AND PeriodeEcriture>=#{self.exefin}#) N1
            ON N.Compte=N1.Compte)            
    """  
        data = self.exec_select(sql)
        for index, row in enumerate(data):
            data[index] = [0.0 if x==None else x for x in row] # None -> 0
        return data

    def maj_solde_comptes(self):

        data = self.calc_solde_comptes()

        logging.info("Mise à jour de la table Comptes")

        for (Numero, Debit, Credit,
             DebitHorsEx, CreditHorsEx, NbEcritures) in data:

            sql = f"""
            UPDATE Comptes 
            SET 
            Debit={Debit},
            Credit={Credit},
            DebitHorsEx={DebitHorsEx},
            CreditHorsEx={CreditHorsEx},
            NbEcritures={NbEcritures}
            WHERE Numero='{Numero}'
            """
            self.exec_insert(sql)


# ///////////////////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////

class QueryDossierAnnuel(object):
    def __init__(self):

        self.cycles = []
        self.cursor = ""
        self.root = ""
    
    def connect(self, chem_base):
        chem_base = os.path.abspath(chem_base)
        self.root = os.path.dirname(chem_base)
        self.chem_base = chem_base.lower()

        constr = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=' + \
            self.chem_base
        try:
            self.conx = pyodbc.connect(constr, autocommit=True)
            logging.info('Ouverture de {}'.format(self.chem_base))
            self.cursor = self.conx.cursor()

        except pyodbc.Error:
            logging.error("erreur requete base {} \n {}".format(
                self.chem_base, sys.exc_info()[1]))
            return False
        except:
            logging.error("erreur ouverture base {} \n {}".format(
                self.chem_base, sys.exc_info()[0]))
            return False
          
        # Table Cycles

        sql = """
            SELECT Code, CodeCycleParam, Libelle, ProchainNumPJ, ListeComptes 
            FROM Cycles
        """
        self.cursor.execute(sql)
        data = self.cursor.fetchall()
        self.cycles = []
        for code, _, libelle, numpj, listecpt in data:
            cycle_items = [code, libelle, numpj]
            cptrng = listecpt.split(";")
            if cptrng[0] == "": #skip if empty
                continue
            liste_plages = []
            if cptrng[0] in ["C", "F"]:
                cycle_items.append(cptrng[0])
            else :
                # on parse les plage de compte de type
                # 610 : 626;628 : 629;650 : 652;656 : 659
                # pour créer une liste
                # de tuple (compte_debut, compte_fin)
                for plage in cptrng:
                    if ":" in plage:
                        deb, fin = plage.split(":")
                        deb = deb.strip()
                        fin = fin.strip()
                        deb = int(deb + ("0" * (8-len(deb))))
                        fin = int(fin + ("9" * (8-len(fin))))
                        liste_plages.append((deb, fin))
                cycle_items.append(liste_plages)
            
            self.cycles.append(cycle_items)

        return self.chem_base.split("\\")[-1]

    def close(self):
        logging.info('fermeture de la base')
        self.conx.commit()
        self.conx.close()


    def _find_code_cycle(self, compte, prefix_frn="0", prefix_cli="9"):
        """
        retourne le code cycle qui correspond au compte
        se refere aux plages indiquées dans la table Cycles
        """
        values = ("", "", "")
        auxil = ""
        if compte.startswith(prefix_frn):
            auxil = "F"
            for code, intitule, last_pj, liste_plage in self.cycles:
                if liste_plage == auxil:
                    values = (code, intitule, last_pj)
                    break
        elif compte.startswith(prefix_cli):
            auxil = "C"
            for code, intitule, last_pj, liste_plage in self.cycles:
                if liste_plage == auxil:
                    values = (code, intitule, last_pj)
                    break
        else:
            compte = int(compte + ("0" * (8-len(compte))))
            for code, intitule, last_pj, liste_plage in self.cycles:
                if len(liste_plage)>1:
                    for deb, fin in liste_plage:
                        if compte in range(deb, fin):
                            values = (code, intitule, last_pj)
                            break
        return values

    def _req_commentaires_indice(self, compte):
        """
        Renvoi numéro d'indice du dernier commentaire de type 3 (PJ)
        de la table CommentairesRevCpt
        """
        sql = """
            SELECT MAX(IndiceLigne) FROM CommentairesRevCpt 
            WHERE NumeroCompte='{}' 
            AND TypeCom=3
        """.format(compte)
        self.cursor.execute(sql)
        indice = self.cursor.fetchall()[0][0]
        if indice == None:
            indice = 1
        return indice

    def cycle_update(self, numPJ, cycle):
        """
        Met à jour la valeur du dernier numéro de pièce dans la table Cycle
        """
        sql = """
            UPDATE Cycles 
            SET ProchainNumPJ={}
            WHERE Code='{}'
        """.format(numPJ, cycle)
        try:
            self.cursor.execute(sql)
        except pyodbc.Error:
            logging.error("erreur insert sur Cycles {}".format(sys.exc_info()[1]))
            logging.debug(sql)
            return False  

        return True

    def commentaire_insert(self, compte, filename):
        """
        Insert nouvelle ligne dans la table CommentairesRevcpt
        Pour l'ajout d'un nouveau doc lié à un compte
        Renvoi True si insert OK
        """
        indice = self._req_commentaires_indice(compte) + 1 
        cycle, intitule, last = self._find_code_cycle(compte)
        if not cycle :
            return ["", ""] 
        last += 1
        nomref = cycle + "P-" + str(last)

        sql = """
            INSERT INTO CommentairesRevCpt
            (
                NumeroCompte, TypeCom, IndiceLigne, NomRef, NomPJ
            )
            VALUES
            (
                '{}', 3, {}, '{}', '{}'
            )
        """.format(compte, indice, nomref, filename)
        try:
            self.cursor.execute(sql)
            self.cycle_update(last, cycle)
        except pyodbc.Error:
            logging.error("erreur insert sur Commentaires {}".format(sys.exc_info()[1]))
            logging.debug(sql)
            return ["", "", ""]  

        return [cycle, intitule, nomref]

    def copy_to_images(self, filepath):
        """
        Méthode pour copier une nouvelle pièce comptable dans le
        dossier images
        Renvoi le basename utilisé
        """
        source_path = os.path.abspath(filepath)
        filename = os.path.basename(source_path)
        images_path = os.path.join(self.root, "Images")
        if not os.path.isdir(images_path):
            os.mkdir(images_path)
        # Si un doc identique est déjà présent dans le dossier
        if os.path.isfile(os.path.join(images_path, filename)):
            filename = doc_rename(filename)
        dest_path = os.path.join(images_path, filename)
        try:
            shutil.copyfile(source_path, dest_path)
        except IOError as e:
            logging.error("Echec copie vers Images : {}".format(e))
            return ""
        return filename

    def ajout_piece(self, compte, filepath):
        filename = self.copy_to_images(filepath)
        # Si la copie du fichier a fonctionné
        # on met à jour la table commentaires
        if filename:
            cycle, intitule, nomref = self.commentaire_insert(compte, filename)
            return (cycle, intitule, nomref, filename)
        
        
             

        



if __name__ == '__main__':

    import pprint
    pp = pprint.PrettyPrinter(indent=4)
    logging.basicConfig(level=logging.DEBUG,
                        format='%(funcName)s\t\t%(levelname)s - %(message)s')


    # cpta = "//srvquadra/qappli/quadra/database/cpta/da209912/Form05/qcompta.mdb"
    # cpta = "//srvquadra/qappli/quadra/database/cpta/ds2099/000175/qcompta.mdb"
    # cpta = "C:/quadra/database/cpta/DC/T00752/qcompta.mdb"
    # da = "C:/quadra/database/cpta/DC/T00752/QDR1812.mdb"
    cpta = "assets/predi_nocent.mdb"

    QC = QueryCompta()
    QC.connect(cpta)
    print(QC.preffrn)
    pp.pprint(QC.comptes)

    # QDA = QueryDossierAnnuel()
    # QDA.connect(da)
    # print(QDA.ajout_piece("40810000", "C:/temp/Leroy.pdf"))
    # print(QDA.ajout_piece("37100000", "C:/temp/Leroy.pdf"))
    # pp.pprint(QDA.cycles)
    # print("7971", QDA.find_code_cycle("7971"))
    # print("9NEWTON", QDA.find_code_cycle("9NEWTON", QC.prefcli))
    # print("9NEWTON", QDA.find_code_cycle("9NEWTON"))
    # print("0NEWTON", QDA.find_code_cycle("0NEWTON"))
    # print(QDA.commentaires_indice("40810000"))
    # print(QDA.commentaires_indice("99999999"))
    # QDA.commentaire_insertpj("40810000", "toto.pdf")
    # QDA.commentaire_insertpj("37100000", "titi.pdf")
    # QDA.close()
    QC.close()
    # print(list_dossier_annuel(cpta))
    # pp.pprint(data["datevalid"])
    # pp.pprint(data["dateclot"])
    # pp.pprint(data["datesortie"])

    # periode = datetime(year=2018, month=1, day=1)
    # o.insert_ecrit("0YYYYYY7", "CA", "000", periode, "ABAQUE", 999.99, 0, "XXXXX", "", "", "")
    # o.insert_ecrit("60110000", "CA", "000", periode, "ABAQUE", 0, 999.99, "XXXXX", "", "", "")

    # o.close()

    # compte = "60110000"

    # print(o.get_affectation_ana(compte))
    # print(o.get_last_uniq())
    # o.insert_compte("42501300")
    # pp.pprint(o.maj_solde_comptes())
    # print(o.get_last_lignefolio("AC", periode))
    # o.maj_central_all()
    # o.maj_central("AC", periode, 0)

    # pp.pprint(QDA.cycles)
    # QDA.find_code_cycle("41")
    # QDA.find_code_cycle("168861")
    # QDA.find_code_cycle("168861")
    # QDA.close()

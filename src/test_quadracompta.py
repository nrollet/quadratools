import pytest
import pyodbc
from random import choice
from datetime import datetime
from quadratools.quadracompta import QueryCompta


# @pytest.fixture
# def cnxn():
#     chem_base = "assets/frozen.mdb"
#     constr = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=' + chem_base
#     conx = pyodbc.connect(constr, autocommit=True)
#     yield conx
#     conx.close()

# @pytest.fixture
# def cursor(cnxn):
#     cursor = cnxn.cursor()
#     yield cursor
#     cnxn.rollback()




def test_maj_centralisateurs():
    """
    Compare une table centralisateur recalculée (Qtest)
    avec une table de la base de référence (Qref)
    """
    Qtest = QueryCompta()
    Qtest.connect("assets/predi_test.mdb")
    Qtest.maj_centralisateurs()
    calc = Qtest.exec_select("SELECT * FROM Centralisateur")

    QRef = QueryCompta()
    QRef.connect("assets/predi_ref.mdb")
    ref = QRef.exec_select("SELECT * FROM Centralisateur")

    assert sorted(calc)==sorted(ref)
    Qtest.close()
    QRef.close()

def test_maj_solde_comptes():
    """
    Compare les soldes des comptes calculés
    avec ceux présents dans la table Comptes
    de la bdd de référence
    """

    Qtest = QueryCompta()
    Qtest.connect("assets/predi_test.mdb")  
    # RAZ des valeurs sur la bdd test
    Qtest.exec_insert("""
        UPDATE Comptes 
        SET Debit=0.0,
        Credit=0.0,
        DebitHorsEx=0.0,
        CreditHorsEx=0.0,
        NbEcritures=0.0"""
        ) 
    # MAJ des soldes
    Qtest.maj_solde_comptes()
    # Collecte table Comptes
    data_test = Qtest.exec_select(
        """
        SELECT Numero, 
        ROUND(Debit, 2), ROUND(Credit, 2), 
        ROUND(DebitHorsEx,2), ROUND(CreditHorsEx, 2), 
        NbEcritures
        FROM Comptes"""
    )   

    QRef = QueryCompta()
    QRef.connect("assets/predi_ref.mdb")
    data_ref = QRef.exec_select(
        """
        SELECT Numero, 
        ROUND(Debit, 2), ROUND(Credit, 2), 
        ROUND(DebitHorsEx,2), ROUND(CreditHorsEx, 2), 
        NbEcritures
        FROM Comptes"""
    )
    assert sorted(data_test)==sorted(data_ref)
    Qtest.close()
    QRef.close()

def test_insert_compte():
    """
    Supprime un compte de la base test
    puis le recréé pour le comparer
    avec la base de référence
    """
    Qtest = QueryCompta()
    Qtest.connect("assets/predi_test.mdb")
    # Sélection d'un compte fourn/general au hasard
    fourn = choice(
        list((x for x in Qtest.plan.keys() if x.startswith("0")))
    )
    gener = choice(
        list((x for x in Qtest.plan.keys() if x.startswith("2")))
    )
    # print("\n", f"fourn:{fourn}", f"gener:{gener}")
    # Suppression
    Qtest.exec_insert(f"""
        DELETE FROM Comptes WHERE Numero='{fourn}' OR Numero='{gener}'
    """)    
    Qtest.insert_compte(fourn)
    Qtest.insert_compte(gener)
    Qtest.maj_solde_comptes()
    data_test = Qtest.exec_select(
        f"""SELECT 
            Numero, Type, TypeCollectif,
            Debit, Credit, DebitHorsEx, CreditHorsEx,
            Collectif,  NbEcritures, MargeTheorique,
            CompteInactif, QuantiteNbEntier
            FROM Comptes WHERE Numero='{fourn}' OR Numero='{gener}'""")


    QRef = QueryCompta()
    QRef.connect("assets/predi_ref.mdb")
    data_ref = QRef.exec_select(
        f"""SELECT 
            Numero, Type, TypeCollectif,
            Debit, Credit, DebitHorsEx, CreditHorsEx,
            Collectif,  NbEcritures, MargeTheorique,
            CompteInactif, QuantiteNbEntier
            FROM Comptes WHERE Numero='{fourn}' OR Numero='{gener}'""")
    print()
    for row in data_test:
        print("TST: ",row)
    for row in data_ref:
        print("REF: ",row)
    assert data_test==data_ref
    Qtest.close()
    QRef.close()  


def test_insert_ecrit():
    """
    Créé deux écritures témoin, une avec date échéance
    une autre avec analytique, compare le résultat
    avec la base de référence
    """
    # Requête commune
    sql_select = """
        SELECT 
        NumUniq, NumeroCompte,
        CodeJournal, Folio,
        LigneFolio, PeriodeEcriture,
        JourEcriture,  Libelle,
        LibelleMemo,  MontantTenuDebit,
        MontantTenuCredit, MontantSaisiDebit,
        MontantSaisiCredit, Quantite,
        NumeroPiece, DateEcheance,
        CodeLettrage, NumLettrage,
        RapproBancaireOk, NoLotEcritures,
        PieceInterne, Etat,
        NumLigne, TypeLigne,
        ModePaiement, CodeBanque,
        Actif, NumEditLettrePaiement,
        ReferenceTire, RIB,
        DomBanque, Centre,
        Nature, PrctRepartition,
        TypeSaisie, ClientOuFrn,
        DateRapproBancaire, JrnRapproBancaire,
        RefImage, MontantAna,
        MilliemesAna, DatePointageAux,
        EcheanceSimple, CentreSimple,
        CodeTva, PeriodiciteDebut,
        PeriodiciteFin, MethodeTVA,
        NumCptOrigine, CodeDevise,
        BonsAPayer, MtDevForce,
        EnLitige, TvaE4DR,
        TvaE4BI, TvaE4Mt,
        Quantite2, NumEcrEco,
        NoLotTrace, CodeLettrageTiers,
        DateDecTVA, NoLotFactor,
        Validee, IBAN,
        BIC, NoLotIs,
        NumMandat, DateSysValidation,
        EcritureNum        
        FROM Ecritures
        WHERE PeriodeEcriture=#2019-05-01#
        AND CodeJournal='AC'
    """
    # Base de référence
    Qref = QueryCompta()
    Qref.connect("assets/balti_ref.mdb")
    data_ref = Qref.exec_select(sql_select)
    # Base de test
    Qtest = QueryCompta()
    Qtest.connect("assets/balti_test.mdb")
    # Suppression des écritures crées
    # par les tests précédents
    Qtest.exec_insert("""
        DELETE * FROM Ecritures 
        WHERE PeriodeEcriture=#2019-05-01#
        AND CodeJournal='AC'
    """)
    ####### 3 lignes + date échéance sur compte fournisseur #######
    Qtest.insert_ecrit(
        compte="0PREV", 
        journal="AC", 
        folio=0, 
        date=datetime(2019, 5, 22),
        libelle="LIGNE AVEC ECHEANCE",
        debit=0.0,
        credit=639.30,
        piece="PIECE1",
        echeance=datetime(2019, 5, 31),
    )
    Qtest.insert_ecrit(
        compte="48861109", 
        journal="AC", 
        folio=0, 
        date=datetime(2019, 5, 22),
        libelle="LIGNE AVEC ECHEANCE",
        debit=532.75,
        credit=0.0,
        piece="PIECE1",
    )
    Qtest.insert_ecrit(
        compte="44566000", 
        journal="AC", 
        folio=0, 
        date=datetime(2019, 5, 22),
        libelle="LIGNE AVEC ECHEANCE",
        debit=106.55,
        credit=0.0,
        piece="PIECE1",
    )
    ####### 3 lignes + analytique sur le compte de charge #######
    Qtest.insert_ecrit(
        compte="0HK", 
        journal="AC", 
        folio=0, 
        date=datetime(2019, 5, 22),
        libelle="LIGNE AVEC ANALYTIQUE",
        debit=0.0,
        credit=7.26,
        piece="PIECE2",
    )
    Qtest.insert_ecrit(
        compte="60630000", 
        journal="AC", 
        folio=0, 
        date=datetime(2019, 5, 22),
        libelle="LIGNE AVEC ANALYTIQUE",
        debit=6.05,
        credit=0.0,
        piece="PIECE2",
        centre="019"
    )
    Qtest.insert_ecrit(
        compte="44566000", 
        journal="AC", 
        folio=0, 
        date=datetime(2019, 5, 22),
        libelle="LIGNE AVEC ANALYTIQUE",
        debit=1.21,
        credit=0.0,
        piece="PIECE2",
    )
    data_test = Qtest.exec_select(sql_select)
    assert sorted(data_ref) == sorted(data_test)

    Qref.close()
    Qtest.close()




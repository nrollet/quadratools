import pytest
import pyodbc
import os
import hashlib
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

def hashfile(path, blocksize = 65536):
    afile = open(path, 'rb')
    hasher = hashlib.md5()
    buf = afile.read(blocksize)
    while len(buf) > 0:
        hasher.update(buf)
        buf = afile.read(blocksize)
    afile.close()
    return hasher.hexdigest()


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
        piece="PIECE2",
    )
    data_test = Qtest.exec_select(sql_select)
    assert sorted(data_ref) == sorted(data_test)

    Qref.close()
    Qtest.close()


def test_ajout_image():
    """
    test la copie d'un fichier vers le dossier images
    Compare le hash MD5 de la source avec la destination
    """
    source_fact = "assets/facture1.pdf"
    hash_ref = hashfile(source_fact)

    Qtest = QueryCompta()
    Qtest.connect("assets/predi_test.mdb")
    dest_fact = Qtest.ajout_image(source_fact)
    hash_test = hashfile(dest_fact)

    assert hash_ref == hash_test
        
    os.remove(dest_fact)
    Qtest.close()

def test_maj_refImage():
    """
    teste la fonction maj_refImage()
    """
    Qref = QueryCompta()
    Qref.connect("assets/predi_ref.mdb")
    imageRef_ref = Qref.exec_select("SELECT refImage FROM Ecritures WHERE NumUniq=8982")[0][0]

    Qtest = QueryCompta()
    Qtest.connect("assets/predi_test.mdb")
    Qtest.exec_insert("UPDATE Ecritures SET refImage='' WHERE NumUniq=8982")
    imageRef_test = Qtest.exec_select("SELECT refImage FROM Ecritures WHERE NumUniq=8982")[0][0]

    print(imageRef_ref, imageRef_test)
    numuniq = 8982
    image_name = "000874_52401.pdf"
    Qtest.maj_refImage(numuniq, image_name)
    imageRef_test = Qtest.exec_select("SELECT refImage FROM Ecritures WHERE NumUniq=8982")[0][0]
    print(imageRef_ref, imageRef_test)

    Qref.close()
    Qtest.close()



if __name__ == "__main__":
    test_maj_refImage()
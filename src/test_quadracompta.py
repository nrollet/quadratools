import pytest
import pyodbc
from random import choice
from quadratools.quadracompta import QueryCompta


@pytest.fixture
def cnxn():
    chem_base = "assets/frozen.mdb"
    constr = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=' + chem_base
    conx = pyodbc.connect(constr, autocommit=True)
    yield conx
    conx.close()

@pytest.fixture
def cursor(cnxn):
    cursor = cnxn.cursor()
    yield cursor
    cnxn.rollback()



def test_rs(cursor):
    sql = """
    SELECT
    RaisonSociale, DebutExercice, FinExercice,
    PeriodeValidee, PeriodeCloturee
    FROM Dossier1
    """
    data = cursor.execute(sql).fetchall()
    assert data[0][0] == "JUSTOM"

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
    avec la abse de référence
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
            CompteInactif, QuantiteNbEntier, DetailCloture, ALettrerAuto
            FROM Comptes WHERE Numero='{fourn}' OR Numero='{gener}'""")


    QRef = QueryCompta()
    QRef.connect("assets/predi_ref.mdb")
    data_ref = QRef.exec_select(
        f"""SELECT 
            Numero, Type, TypeCollectif,
            Debit, Credit, DebitHorsEx, CreditHorsEx,
            Collectif,  NbEcritures, MargeTheorique,
            CompteInactif, QuantiteNbEntier, DetailCloture, ALettrerAuto
            FROM Comptes WHERE Numero='{fourn}' OR Numero='{gener}'""")
    print()
    for row in data_test:
        print("TST: ",row)
    for row in data_ref:
        print("REF: ",row)
    assert data_test==data_ref
    Qtest.close()
    QRef.close()  





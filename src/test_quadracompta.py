import pytest
import pyodbc
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
    print(data)
    assert data[0][0] == "JUSTOM"

def test_maj_centralisateurs():
    """
    Compare une table centralisateur recalculée (QCalc)
    avec une table de la base de référence (Qref)
    """
    QCalc = QueryCompta()
    QCalc.connect("assets/predi_nocent.mdb")
    QCalc.maj_centralisateurs()
    calc = QCalc.exec_select("SELECT * FROM Centralisateur")

    QRef = QueryCompta()
    QRef.connect("assets/predi.mdb")
    ref = QRef.exec_select("SELECT * FROM Centralisateur")

    assert calc==ref
    QCalc.close()
    QRef.close()

def test_maj_solde_comptes():
    """
    Compare les soldes des comptes calculés
    avec ceux présents dans la table Comptes
    de la bdd de référence
    """

    QCalc = QueryCompta()
    QCalc.connect("assets/predi_nocent.mdb")  
    QCalc.exec_select("""
        UPDATE Comptes 
        SET Debit=0.0,
        Credit=0.0,
        DebitHorsEx=0.0,
        CreditHorsEx=0.0,
        NbEcritures=0.0"""
        ) # RAZ des valeurs sur la bdd test
    QCalc.maj_solde_comptes()
    calc = QCalc.exec_select("""
        SELECT Comptes,
        Debit, Credit,
        DebitHorsEx, CreditHorsEx,
        NbEcritures
        FROM Comptes"""
    )        
    QRef = QueryCompta()
    QRef.connect("assets/predi.mdb")
    ref = QRef.exec_select(
        """
        SELECT Comptes,
        Debit, Credit,
        DebitHorsEx, CreditHorsEx,
        NbEcritures
        FROM Comptes"""
    )
    assert calc==ref
    QCalc.close()
    QRef.close()





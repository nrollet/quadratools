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

def test_pass():
    assert True, "dummy sample test"


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


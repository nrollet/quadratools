import pytest
from quadratools.quadracompta import QueryCompta

# @pytest.fixture
# def db


def test_pass():
    assert True, "dummy sample test"


def test_connect():
    dbpath = "assets/frozen.mdb"
    Q = QueryCompta()
    Q.connect(dbpath)
    assert Q.rs == "JUSTOM"
    Q.close()


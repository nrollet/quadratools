import logging
import pprint
from collections import namedtuple
from quadratools.quadracompta import QueryCompta

pp = pprint.PrettyPrinter(indent=4)

logging.basicConfig(
    level=logging.DEBUG, format="%(module)s-%(funcName)s\t\t%(levelname)s - %(message)s"
)

Ecriture = namedtuple(
    "ecriture",
    [
        "compte",
        "journal",
        "folio",
        "date",
        "libelle",
        "debit",
        "credit",
        "piece",
        "centre",
    ],
)

dbpath = "assets/predi_test.mdb"

Q = QueryCompta()
Q.connect(dbpath)

data = Q.exec_select(
    """
    SELECT * FROM Ecritures
    WHERE
    NumUniq=(SELECT MAX(NumUniq) FROM Ecritures)
    """
)[0]
print(data)
uniq = data[0]
ecr = Ecriture(
    data[1], data[2], data[3], data[5], data[8], data[10], data[11], data[16], data[35]
)

Q.exec_insert(
    f"""
    DELETE * FROM Ecritures WHERE NumUniq=
    """
)

print(*ecr)
Q.close()

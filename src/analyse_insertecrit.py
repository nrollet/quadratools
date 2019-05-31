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

dbpath = "//srvquadra/qappli/quadra/database/cpta/dc/form05/qcompta.mdb"

Q = QueryCompta()
Q.connect(dbpath)

# data = Q.exec_select(
#     """
#     SELECT EcheanceSimple FROM Ecritures WHERE NumUniq=1
#     """
# )
# print(data)
data = Q.exec_select(
    """
    SELECT * FROM Ecritures
    WHERE
    NumUniq=(SELECT MAX(NumUniq) FROM Ecritures)
    """
)[0]
# print(";".join(list(data)))
uniq = data[0]
ecr = Ecriture(
    data[1], data[2], data[3], data[5], data[8], data[10], data[11], data[16], data[35]
)
# # lst = list(*ecr)
# # print(";".join(lst))
sql = f"""
    DELETE * FROM Ecritures WHERE NumUniq={uniq}
"""
# # print(sql)
Q.exec_insert(sql)

print(*ecr)
uid = Q.insert_ecrit(
    ecr.compte,
    ecr.journal,
    ecr.folio,
    ecr.date,
    ecr.libelle+"_Y",
    ecr.debit,
    ecr.credit,
    ecr.piece,
)
print(uid)

Q.maj_centralisateurs()
Q.maj_solde_comptes()

Q.close()

import logging
import pprint
from collections import namedtuple
from quadratools.quadracompta import QueryCompta

pp = pprint.PrettyPrinter(indent=4)

logging.basicConfig(
    level=logging.DEBUG, format="%(module)s-%(funcName)s\t\t%(levelname)s - %(message)s"
)

Compte = namedtuple(
    "compte", ["numero", "debit", "credit", "debitHex", "creditHex", "nbecr"]
)

dbpath = "assets/predi.mdb"
# dbpath = "//srvquadra/qappli/quadra/database/cpta/dc/DEMO/qcompta.mdb"

Q = QueryCompta()
Q.connect(dbpath)

data = Q.calc_solde_comptes()

calc_list = []
for row in data:
    calc_list.append(Compte(*row))

sql = f"""
    SELECT Numero, 
    ROUND(Debit, 2), ROUND(Credit, 2), 
    ROUND(DebitHorsEx,2), ROUND(CreditHorsEx, 2), 
    NbEcritures
    FROM Comptes
"""
data = Q.exec_select(sql)
ref_list = []
for row in data:
    ref_list.append(Compte(*row))

for index, row in enumerate(calc_list):
    for item in ref_list:
        if row.numero == item.numero:
            if row != item:
                print(row, "\n\t", item)

Q.close()

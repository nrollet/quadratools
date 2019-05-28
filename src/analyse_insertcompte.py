import logging
import pprint
from random import choice
from collections import namedtuple
from quadratools.quadracompta import QueryCompta

pp = pprint.PrettyPrinter(indent=4)

logging.basicConfig(
    level=logging.DEBUG, format="%(module)s-%(funcName)s\t\t%(levelname)s - %(message)s"
)

dbpath = "assets/predi_test.mdb"

Q = QueryCompta()
Q.connect(dbpath)

# fourn = list(filter(lambda x: x.startswith("0"), Q.plan.keys()))
# fourn = choice(list((x for x in Q.plan.keys() if x.startswith("0"))))
# print(fourn)
Q.exec_insert(
    f"""
        DELETE FROM Comptes WHERE Numero='0LEVILLA'
    """
)

Q.insert_compte("0LEVILLA")
Q.maj_solde_comptes()
data_test = Q.exec_select(
        f"""SELECT * FROM Comptes WHERE Numero='0LEVILLA'""")
pp.pprint(data_test)

Q.close()

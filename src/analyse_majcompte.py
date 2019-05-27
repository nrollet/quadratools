import logging
import pprint
from collections import namedtuple
from quadratools.quadracompta import QueryCompta

pp = pprint.PrettyPrinter(indent=4)

logging.basicConfig(level=logging.DEBUG,
                    format='%(module)s-%(funcName)s\t\t%(levelname)s - %(message)s')

Compte = namedtuple("compte", ["numero", "debit", "credit", "debitHex", "creditHex", "nbecr"])

dbpath = "assets/predi.mdb"

Q = QueryCompta()
Q.connect(dbpath)

sql = f"""
SELECT NB.NumeroCompte,
    N.debit, N.credit,
    N1.debit, N1.credit,
    NB.NbEcritures
FROM ((
    SELECT NumeroCompte, COUNT(*) AS NbEcritures
    FROM Ecritures
    WHERE TypeLigne='E'
    GROUP BY NumeroCompte) NB
    LEFT JOIN(
    SELECT NumeroCompte, SUM(MontantTenuDebit) AS debit, SUM(MontantTenuCredit) as credit
    FROM Ecritures
    WHERE TypeLigne='E'
    GROUP BY NumeroCompte) N
    ON NB.NumeroCompte=N.Numerocompte)
    LEFT JOIN (
    SELECT NumeroCompte, SUM(MontantTenuDebit) AS debit, SUM(MontantTenuCredit) as credit
    FROM Ecritures
    WHERE PeriodeEcriture>=#{Q.exefin}# 
    AND TypeLigne='E'
    GROUP BY NumeroCompte) N1
    ON NB.NumeroCompte=N1.NumeroCompte
    """ 
calc = Q.exec_select(sql)
calc_list = []
for row in calc:
    ntlist.append(Compte(*row))

sql = f"""
    SELECT Numero, Debit, Credit, DebitHorsEx, CreditHorsEx, NbEcritures
    FROM Comptes
"""
ref = Q.exec_select(sql)

pp.pprint(ntlist)

Q.close()

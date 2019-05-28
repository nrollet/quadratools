import pprint
import datetime 
import logging
# from quadratools.sqlstr import requete_central
from quadratools.quadracompta import QueryCompta

pp = pprint.PrettyPrinter(indent=4)

logging.basicConfig(level=logging.DEBUG,
                    format='%(module)s-%(funcName)s\t\t%(levelname)s - %(message)s')

dbpath = "assets/predi_test.mdb"
# dbpath = "assets/ajusted.mdb"
# dbpath = "//srvquadra/qappli/quadra/database/cpta/dc/000874/qcompta.mdb"

Q = QueryCompta()
Q.connect(dbpath)

central_ref = Q.exec_select("SELECT * FROM Centralisateur")
central = Q.calc_centralisateurs()

for i, row in enumerate(central):
    ## datetime changée en string
    row = [x.strftime("%d/%m/%Y") if isinstance(x, datetime.datetime) else x for x in row]
    ## On convertit tout en int pour éviter les pbm d'arrondis
    central[i] = [int(x) if isinstance(x, float) else x for x in row]

for i, row in enumerate(central_ref):
    ## datetime changée en string sauf si == ''
    row = [x.strftime("%d/%m/%Y") if isinstance(x, datetime.datetime) else x for x in row]
    ## les à-nouveaux vont retourner une date 1899, on la change
    row = ["" if x=="30/12/1899" else x for x in row]
    ## On convertit tout en int pour éviter les pbm d'arrondis
    central_ref[i] = [int(x) if isinstance(x, float) else x for x in row]

# for item in central:
#     print(item)
print("-=-=-=-=- Lignes divergentes issues de la table ecritures -=-=-=-=-")
for index, row in enumerate(sorted(central)):
    if row not in central_ref:
        print(index, row)

print("-=-=-=-=- Lignes divergentes issues de la table centralisateurs -=-=-=-=-")
for index, row in enumerate(sorted(central_ref)):
    if row not in central:
        print(index, row)

# Q.maj_centralisateurs()

Q.close()

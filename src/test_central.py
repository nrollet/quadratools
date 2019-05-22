import pprint
import datetime 
from quadratools.sqlstr import SqlStr
from quadratools.quadracompta import QueryCompta

pp = pprint.PrettyPrinter(indent=4)

dbpath = "assets/frozen.mdb"

Q = QueryCompta()
Q.connect(dbpath)

central_ref = Q.exec_select("SELECT * FROM Centralisateur")
# central_ref = Q.exec_select(SqlStr.SEL_CENTRAL)
# pp.pprint(central_ref)

central = Q.calc_centralisateurs()

for i, row in enumerate(central):
    ## datetime changée en string
    row = [x.strftime("%d/%m/%Y") if isinstance(x, datetime.datetime) else x for x in row]
    ## Remplacement des valeurs 'None' en 0.0
    row = [0.0 if x == None else x for x in row]
    ## On convertit tout en int pour éviter les pbm d'arrondis
    central[i] = [int(x) if isinstance(x, float) else x for x in row]

for i, row in enumerate(central_ref):
    ## datetime changée en string sauf si == ''
    row = [x.strftime("%d/%m/%Y") if isinstance(x, datetime.datetime) else x for x in row]
    ## les à-nouveaux vont retourner une date 1899, on la change
    row = ["" if x=="30/12/1899" else x for x in row]
    ## Remplacement des valeurs 'None' en 0.0
    row = [0.0 if x == None else x for x in row]
    ## On convertit tout en int pour éviter les pbm d'arrondis
    central_ref[i] = [int(x) if isinstance(x, float) else x for x in row]

print("-=-=-=-=- Lignes divergentes issues de la table ecritures -=-=-=-=-")
for row in central:
    if row not in central_ref:
        print(row)

print("-=-=-=-=- Lignes divergentes issues de la table centralisateurs -=-=-=-=-")
for row in central_ref:
    if row not in central:
        print(row)

# tot_res_debit = 0.0
# tot_res_credit = 0.0

# ## Remplacement des valeurs 'None' en 0.0
# for i, row in enumerate(central):
#     central[i] = [0.0 if x == None else x for x in row]
# ## On convertit tout en int pour éviter les pbm d'arrondis)
# for i, row in enumerate(central):
#     central[i] = [int(x) if isinstance(x, float) else x for x in row]
# for i, row in enumerate(central):
#     central[i] = [x.strftime("%d/%m/%Y") if isinstance(x, datetime.datetime) else x for x in row]
#     # central[i] = [x.strftime("%d/%m/%Y") if type(x)==datetime.datetime else x for x in row]
# for i, row in enumerate(central_ref):
#     central_ref[i] = [int(x) if isinstance(x, float) else x for x in row]





Q.close()

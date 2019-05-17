import pprint
from quadratools.quadracompta import QueryCompta

pp = pprint.PrettyPrinter(indent=4)

dbpath = "assets/frozen.mdb"

Q = QueryCompta()
Q.connect(dbpath)

central_ref = Q.exec_select("SELECT * FROM Centralisateur")

central = Q.calc_centralisateurs()

tot_res_debit = 0.0
tot_res_credit = 0.0

# Remplacement des valeurs 'None' en 0.0
for i, row in enumerate(central):
    central[i] = [0.0 if x == None else x for x in row]
# On convertit tout en int pour éviter les pbm d'arrondis)
for i, row in enumerate(central):
    central[i] = [int(x) if isinstance(x, float) else x for x in row]
for i, row in enumerate(central_ref):
    central_ref[i] = [int(x) if isinstance(x, float) else x for x in row]



print("-=-=-=-=- manquantes ou erronées dans ecritures -=-=-=-=-")
for row in central:
    if row not in central_ref:
        print(row)

print("-=-=-=-=- manquantes ou erronées dans centralisateur -=-=-=-=-")
for row in central_ref:
    if row not in central:
        print(row)


Q.close()

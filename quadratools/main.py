import logging
import pprint
from qpaietools import QueryPaie


pp = pprint.PrettyPrinter(indent=4)
logging.basicConfig(
    level=logging.DEBUG, format="%(funcName)s\t\t%(levelname)s - %(message)s"
)

ipl = "C:/Users/nicolas/Documents/Pydio/mono.ipl"
Q = QueryPaie(ipl, "FORM05")
Q.connect()

pp.pprint(Q.param_dossier())
# pp.pprint(Q.employes())
# pp.pprint(Q.dump_table("employes"))
Q.generate_provcp_report()

Q.disconnect()

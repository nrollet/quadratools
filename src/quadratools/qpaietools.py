import logging
import pyodbc
from datetime import datetime, timedelta
from quadraenv import QuadraSetEnv


def getcol(column_descr):

    col_list = list()

    for col_name, col_type, _, _, _, _, _ in column_descr:
        if col_type == type(1):
            datatype = "INT"
        elif col_type == type(True):
            datatype = "TINYINT"
        elif col_type == type(1.0):
            datatype = "FLOAT"
        elif col_type == type("A"):
            datatype = "VARCHAR(100)"
        elif col_type == type(datetime.now()):
            datatype = "DATETIME"
        col_list.append((col_name, datatype))
    return col_list


class QueryPaie(object):
    """
    Submit queries to Quadra Paie database in MS Access format
    """

    def __init__(self, ipl_file, code_dossier):

        self.conx = ""
        self.code_dossier = code_dossier
        self.param = {}
        self.periodepaie = ""
        self.databull = []
        self.dataemp = []
        self.provcp = []
        self.qpaie = {"etabl": [], "periode": ""}

        qenv = QuadraSetEnv(ipl_file)
        self.mdb_path = qenv.make_db_path(self.code_dossier, "PAIE")
        logging.info("target db is {}".format(self.mdb_path))

    def connect(self):

        constr = (
            "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" + self.mdb_path
        )
        try:
            self.conx = pyodbc.connect(constr, autocommit=True)
            self.cursor = self.conx.cursor()
            logging.info("Connection OK")
        except pyodbc.Error as msg:
            logging.error(msg)
            return False
        return True

    def disconnect(self):
        self.conx.commit()
        self.conx.close()

    def param_dossier(self):
        """requête pour les paramètres des établissements"""
        sql = "SELECT CodeEtablissement, RaisonSociale FROM Etablissements"
        self.cursor.execute(sql)
        for etab, rs in self.cursor.fetchall():
            # self.param.update({"etab": {"rs": rs, "code": etab}})
            logging.info("Etablissement {} : {}".format(etab, rs))
            self.qpaie["etabl"].append([etab, rs])

        sql = """SELECT PeriodePaie FROM ConstantesEntreprise"""
        self.cursor.execute(sql)
        self.periodepaie = self.cursor.fetchall()[0][0]
        logging.info("Periode en cours : {}".format(self.periodepaie.strftime("%m/%Y")))
        self.qpaie.update({"periode": self.periodepaie})

        # return self.qpaie

    def employes(self):

        # self.qpaie.

        sql = """
            SELECT EM.Numero, EM.NomNaissance,
            EM.NomMarital, EM.Prenom,
            EM.DateEntree1, EM.DateEntree2,
            EM.DateSortie1, EM.DateSortie2
            FROM Employes EM
            INNER JOIN CriteresLibres CL
            ON EM.Numero=CL.NumeroEmploye
            """

        self.cursor.execute(sql)
        data = self.cursor.fetchall()
        # columns = getcol(self.cursor.description)
        # print(columns)

        for (
            Numero,
            NomNaissance,
            NomMarital,
            Prenom,
            DateEntree1,
            DateEntree2,
            DateSortie1,
            DateSortie2,
        ) in data:
            numero = Numero
            # numero = Numero.lstrip()
            nom = NomNaissance
            if NomMarital:
                nom = NomMarital
            nom = nom + " " + Prenom
            entree = DateEntree1
            if DateEntree2 > DateEntree1:
                entree = DateEntree2
            sortie = DateSortie1
            if DateSortie1 and DateEntree2 == "":
                sortie = DateSortie1
            elif DateSortie2:
                sortie = DateSortie2
            if entree == datetime(1899, 12, 30, 0, 0):
                entree = None
            if sortie == datetime(1899, 12, 30, 0, 0):
                sortie = None

            self.dataemp.append([numero, nom, entree, sortie])
        column_names = (
            ("matricule", "VARCHAR(100)"),
            ("nom", "VARCHAR(100)"),
            ("entree", "DATETIME"),
            ("sortie", "DATETIME"),
        )
        self.qpaie.update(
            {
                "employes" :
                {
                    "columns": column_names,
                    "rows" : self.dataemp
                }
            }
        )

        return self.qpaie

    # def dump_table(self, table_name):
    #     """requête générique pour récupérer une table entière
    #        au format JSON
    #     """
    #     sql = "SELECT * FROM {}".format(table_name)

    #     rdic = {}

    #     self.cursor.execute(sql)
    #     rows = self.cursor.fetchall()
    #     columns = [item[0] for item in self.cursor.description]

    #     for row in rows:
    #         employe = row[0]
    #         rdic.setdefault(employe, {})
    #         for i in range(columns[1:]) :
    #             rdic[employe].update({column: })


        # columns = getcol(self.cursor.description)

        # result_dic.update(
        #     {
        #         table_name :
        #         {
        #             "columns": columns,
        #             "rows" : data
        #         }
        #     }
        # )
        # return result_dic

    # def generate_provcp_report(self, periode):
    # def generate_provcp_report(self):

    #     # Calc date période précédente
    #     # periode = periode.replace(day=1)
    #     # periode_1 = periode - timedelta(days=1)
    #     # periode_1 = periode_1.replace(day=1)

    #     print(self.dump_table("bulletins"))



    def bulletins_provcp(self):
        """
        Collecte la provision de cp du mois des employes
        sur la période choisie
        """
        # Calc date période précédente
        # periode = periode.replace(day=1)
        # periode_1 = periode - timedelta(days=1)
        # periode_1 = periode_1.replace(day=1)

        sql = """
            SELECT N.Periode,
            N.NumeroEmploye,
            E.NomNaissance & ' ' & E.Prenom,
            N.BrutCP,
            N.NbJDus_1,
            N.NbJPris_1,
            N.ProvCP_1,
            N.CumMtCpPris_1,
            N.ProvCp,
            N.NbJDus,
            N.NbJPris,
            N.CumMtCpPris,
            N.DroitCP,
            N.MtJourneeCP,
            N.NbJourCPPris
            FROM Bulletins N
            LEFT JOIN Employes E
            ON N.NumeroEmploye=E.Numero
            ORDER BY N.Periode, N.NumeroEmploye
        """

        self.cursor.execute(sql)
        rows = self.cursor.fetchall()

        table_schema = [
                ("Période", "DATETIME"),
                ("Numéro", "VARCHAR"),
                ("nom", "VARCHAR"),
                ("JoursDusN_1", "FLOAT"),
                ("JoursPrisN_1", "FLOAT"),
                ("SoldeJoursN_1", "FLOAT"),
                ("ProvisionN_1", "FLOAT"),
                ("MontantPrisN_1", "FLOAT"),
                ("SoldeProvisionN_1", "FLOAT"),
                ("JoursDusN", "FLOAT"),
                ("JoursPrisN", "FLOAT"),
                ("SoldeJoursN", "FLOAT"),
                ("ProvisionN", "FLOAT"),
                ("MontantPrisN", "FLOAT"),
                ("SoldeProvisionN", "FLOAT"),
                ("TotalProvision", "FLOAT")
        ]

        # en-tête de colonnes pour la première ligne
        csv_list = [[item[0] for item in table_schema]]

        for (
            Periode, Employe,
            Nom, BrutCP,
            NbJDus_1,
            NbJPris_1,
            ProvCP_1,
            CumMtCpPris_1,
            ProvCp,
            NbJDus,
            NbJPris,
            CumMtCpPris,
            DroitCP, MtJourneeCP, NbJourCPPris
        ) in rows:

            # Solde jours N-1
            solde_jours_1 = NbJDus_1 - NbJPris_1   
            # Si jours posés > au solde n-1 congés pris sur n-1 == solde_jours_1
            if NbJourCPPris > solde_jours_1:
                conges_pris_1 = solde_jours_1
                NbJourCPPris = NbJourCPPris - conges_pris_1
            else : 
                conges_pris_1 = NbJourCPPris
                NbJourCPPris = 0
            # s'il reste des jours sur n-1, le montant des jours 
            # pris n-1 va augmenter du nombre de jours qu'a posé le salarié sur n-1
            if solde_jours_1 > 0:
                montant_pris_1 = CumMtCpPris_1 + (MtJourneeCP * conges_pris_1)
            else : 
                montant_pris_1 = CumMtCpPris_1
            
            # calcul du nouveau solde n-1
            NbJPris_1 = NbJPris_1 + conges_pris_1
            solde_jours_1 = NbJDus_1 - NbJPris_1  
            
            # Solde provision N-1
            if NbJDus_1 > 0:
                solde_provision_1 = ProvCP_1 - (CumMtCpPris_1 + MtJourneeCP)
            else:
                solde_provision_1 = 0.0
            # Jours dus N
            jours_du_n = NbJDus + DroitCP
            # Jours pris N
            if (NbJDus_1 - NbJPris_1) == 0:
                jours_pris_n = NbJPris + NbJourCPPris
            else:
                jours_pris_n = NbJPris
            # Solde jours N
            solde_jours_n = jours_du_n - jours_pris_n
            # Provisions N
            provisions_n = (BrutCP * 0.1) + ProvCp
            # Montant pris N
            if jours_pris_n > 0:
                montant_pris_n = CumMtCpPris + MtJourneeCP
            else:
                montant_pris_n = CumMtCpPris
            # Solde provisions
            solde_provision = provisions_n - montant_pris_n
            # Total Provisions
            total_provision = solde_provision + solde_provision_1

            csv_list.append(
                [
                    Periode.strftime("%d/%m/%Y"),
                    Employe,
                    Nom,
                    f(NbJDus_1),
                    f(NbJPris_1),
                    f(solde_jours_1),
                    f(ProvCP_1),
                    f(montant_pris_1),
                    f(solde_provision_1),
                    f(jours_du_n),
                    f(jours_pris_n),
                    f(solde_jours_n),
                    f(provisions_n),
                    f(montant_pris_n),
                    f(solde_provision),
                    f(total_provision)
                ]
            )
        return csv_list



if __name__ == "__main__":

    # import locale
    import pprint
    import logging

    def f(d):
        return str(round(d, 2)).replace(".", ",")

    pp = pprint.PrettyPrinter(indent=4)
    logging.basicConfig(
        level=logging.DEBUG, format="%(funcName)s\t\t%(levelname)s - %(message)s"
    )
    ipl = "C:/Users/nicolas/Documents/Pydio/mono.ipl"
    o = QueryPaie(ipl, "FORM05")
    if o.connect():
        logging.info("success")

    data = o.bulletins_provcp()

    # pp.pprint(data)

    with open("provision.csv", "w", encoding="cp1252") as csv:
        for row in data:
            print(row)
            csv.write(";".join(row) + "\n")
    

    o.disconnect()
    # pp.pprint(data)

    # out_file = open("provision.csv", "w")

    # out_file.write(
    #     ";".join(
    #         [
    #             "Période",
    #             "Numéro",
    #             "nom",
    #             "Jours dus N-1",
    #             "Jours pris N-1",
    #             "Solde jours N-1",
    #             "Provision N-1",
    #             "Montant pris N-1",
    #             "Solde provision N-1",
    #             "Jours dus N",
    #             "Jours pris N",
    #             "Solde jours N",
    #             "Provision N",
    #             "Montant pris N",
    #             "Solde provision N",
    #             "Total provision\n",
    #         ]
    #     )
    # )

    # for (
    #     Periode,
    #     Employe,
    #     Nom,
    #     BrutCP,
    #     NbJDus_1,
    #     NbJPris_1,
    #     ProvCP_1,
    #     CumMtCpPris_1,
    #     ProvCp,
    #     NbJDus,
    #     NbJPris,
    #     CumMtCpPris,
    #     DroitCP,
    #     MtJourneeCP,
    #     NbJourCPPris
    # ) in data:

    #     # Solde jours N-1
    #     solde_jours_1 = NbJDus_1 - NbJPris_1   
    #     # Si jours posés > au solde n-1 congés pris sur n-1 == solde_jours_1
    #     if NbJourCPPris > solde_jours_1:
    #         conges_pris_1 = solde_jours_1
    #         NbJourCPPris = NbJourCPPris - conges_pris_1
    #     else : 
    #         conges_pris_1 = NbJourCPPris
    #         NbJourCPPris = 0
    #     # s'il reste des jours sur n-1, le montant des jours 
    #     # pris n-1 va augmenter du nombre de jours qu'a posé le salarié sur n-1
    #     if solde_jours_1 > 0:
    #         montant_pris_1 = CumMtCpPris_1 + (MtJourneeCP * conges_pris_1)
    #     else : 
    #         montant_pris_1 = 0.0
        
    #     # calcul du nouveau solde n-1
    #     NbJPris_1 = NbJPris_1 + conges_pris_1
        
    #     # print("{} {} {} {}".format(Employe, solde_jours_1, NbJourCPPris, CumMtCpPris_1) )
    #     # Solde provision N-1
    #     if NbJDus_1 > 0:
    #         solde_provision_1 = ProvCP_1 - (CumMtCpPris_1 + MtJourneeCP)
    #     else:
    #         solde_provision_1 = 0.0
    #     # Jours dus N
    #     jours_du_n = NbJDus + DroitCP
    #     # Jours pris N
    #     if (NbJDus_1 - NbJPris_1) == 0:
    #         jours_pris_n = NbJPris + NbJourCPPris
    #     else:
    #         jours_pris_n = NbJPris
    #     # Solde jours N
    #     solde_jours_n = jours_du_n - jours_pris_n
    #     # Provisions N
    #     provisions_n = (BrutCP * 0.1) + ProvCp
    #     # Montant pris N
    #     if jours_pris_n > 0:
    #         montant_pris_n = CumMtCpPris + MtJourneeCP
    #     else:
    #         montant_pris_n = CumMtCpPris
    #     # Solde provisions
    #     solde_provision = provisions_n - montant_pris_n
    #     # Total Provisions
    #     total_provision = solde_provision + solde_provision_1

    #     row_list = [
    #         Periode.strftime("%d/%m/%Y"),
    #         Employe,
    #         Nom,
    #         f(NbJDus_1),
    #         f(NbJPris_1),
    #         f(solde_jours_1),
    #         f(ProvCP_1),
    #         f(montant_pris_1),
    #         f(solde_provision_1),
    #         f(jours_du_n),
    #         f(jours_pris_n),
    #         f(solde_jours_n),
    #         f(provisions_n),
    #         f(montant_pris_n),
    #         f(solde_provision),
    #         f(total_provision)
    #     ]
    #     out_file.write(";".join(row_list))
    #     out_file.write("\n")

    # out_file.close()

    # o.disconnect()

    # o = QueryPaie()
    # # data, columns = o.employes()

    # # pp.pprint(data)
    # # pp.pprint(columns)
    # data, columns = o.bulletins()
    # pp.pprint(columns)
    # # tst = o.cursor.description[0][1]
    # # print(type(tst))
    # # if tst == <class 'str'>:
    # #     print("yes")


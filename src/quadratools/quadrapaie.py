#!/usr/bin/env python3
# -*-coding:Utf-8 -*

import pyodbc
import logging
import sys
import os
import pprint
from collections import namedtuple
from datetime import datetime, timedelta

def add_one_month(dt0):

    dt1 = dt0.replace(day=1)
    dt2 = dt1 + timedelta(days=32)
    dt3 = dt2.replace(day=1)
    return dt3



class QueryPaie(object):

    def __init__(self):

        self.conx = ""
        self.db_path = ""
        self.employes = []
           
    def connect(self, db_path):
        """
        Connexion à la bdd quadra paie
        db_path : chemin complet vers le fichier qpaie.mdb
        """
        self.db_path = db_path.lower()
        if not os.path.isfile(self.db_path):
            logging.error(f"Fichier mdb introuvable : {self.db_path}")
            return False

        constr = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=' + \
            self.db_path
        try:
            self.conx = pyodbc.connect(constr, autocommit=True)
            logging.info('Ouverture de {}'.format(self.db_path))
            self.cursor = self.conx.cursor()

        except pyodbc.Error:
            logging.error("erreur requete base {} \n {}".format(
                self.db_path, sys.exc_info()[1]))
        except:
            logging.error("erreur ouverture base {} \n {}".format(
                self.db_path, sys.exc_info()[0]))

    def close(self):
        logging.info('Fermeture de la base')
        self.conx.commit()
        self.conx.close()

    def exec_select(self, sql_string):
        """
        Pour exécuter une requete sql passée en argument
        """
        data = ""
        try:
            data = self.cursor.execute(sql_string).fetchall()
        except pyodbc.Error:
            logging.error("erreur requete base {} \n {}".format(
                self.db_path, sys.exc_info()[1]))
            logging.error("Requete SQL : {}".format(sql_string))                
        except:
            logging.error("erreur ouverture base {} \n {}".format(
                self.db_path, sys.exc_info()[0]))
            logging.error("Requete SQL : {}".format(sql_string))
        return data    

    def Employes(self):

        sql = """
        SELECT 
        EM.Numero, 
        EM.NomNaissance, 
        EM.NomMarital, 
        EM.Prenom, 
        CL.Texte1,
        EM.DateEntree1,
        EM.DateSortie1,
        EM.DateEntree2,
        EM.DateSortie2
        FROM Employes EM 
        INNER JOIN CriteresLibres CL  
        ON EM.Numero=CL.NumeroEmploye         
        """

        data = self.exec_select(sql)

        Employe = namedtuple("Employe",["numqdr", "nummcd", "nom", "prenom", "entree", "sortie"])

        for numero, nom, nommar, prenom, text1, entree1, sortie1, entree2, sortie2 in data:
            if nommar:
                nom =  nommar
            entree = entree1
            sortie = datetime(1899, 12, 30, 0, 0)
            nodate = datetime(1899, 12, 30, 0, 0)
            if entree2 > entree1 : 
                entree = entree2
            elif entree2 >= sortie1 and entree2 != nodate:
                entree = entree2
            if sortie1 >= entree1 and entree2 == nodate:
                sortie = sortie1
            elif sortie2 >= entree2 and sortie2 != nodate: 
                print(nom)
                sortie = sortie2

            employe = Employe(  
                numero.replace(" ", "").zfill(6),
                text1.zfill(6),
                nom,
                prenom,
                entree,
                sortie)

            self.employes.append(employe)

        return self.employes


    def SalariesPresents(self, periode):
        # on ajoute un mois pour la requête
        for employes in self.employes:
            if (
                employe.entree <= periode and
                employe.
            )
            

        

    def creaDict (self) :

        if self.dataEmp :

            jsdata = {}
            cle = ""

            for (numero, 
                 nomnaiss, 
                 nommarit, 
                 prenom, 
                 matrmcdo) in self.dataEmp :

                if matrmcdo :
                    cle = matrmcdo.zfill(6)
                else :
                    cle = numero.replace(' ','').zfill(6)

                subdir = {'NomNaiss':'', 'NomMarit':'', 'Prenom':'', 'matrMcdo':'', 'absDsn':''}

                jsdata.setdefault(cle, subdir)

                jsdata[cle]['NomNaiss'] = nomnaiss
                jsdata[cle]['NomMarit'] = nommarit
                jsdata[cle]['Prenom'] = prenom
                jsdata[cle]['matrQdr'] = numero.replace(' ','')
                #jsdata[item[0]]['absDsn'] = item[]

        else :
            logging.warning("La requête SQL ne retourne aucune donnée")
        
        logging.debug(jsdata)
        return jsdata

class QueryPaieSalariesNoPer(object):

    def __init__(self, chem_base):

        constr = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=' + chem_base
        
            # Requ?te sur la db ListeEmployes
        sSQL = ("SELECT "
                    "EM.Numero, "
                    "EM.NomNaissance, "
                    "EM.NomMarital, "
                    "EM.Prenom, "
                    "CL.Texte1 "
                    "FROM Employes EM "
                    "INNER JOIN CriteresLibres CL " 
                    "ON EM.Numero=CL.NumeroEmploye ")

        logging.debug(sSQL)
        # Ex?cution requ?te SQL 
        self.dataEmp = []
        try :
            conn = pyodbc.connect(constr, autocommit=True)
            logging.info('Ouverture de {}'.format(chem_base))
            cur = conn.cursor()

            cur.execute(sSQL)
            self.dataEmp = list(cur)

            conn.close

            # on el?ve les espaces du matricule
            #for item in self.dataEmp :
            #    item[0] = item[0].replace(' ', '')

        except pyodbc.Error :
            print("erreur requete base {} \n {}".format(chem_base,sys.exc_info()[1]))
        except :
            print("erreur ouverture base {} \n {}".format(chem_base,sys.exc_info()[0]))

        logging.debug (self.dataEmp)

        

    def creaDict (self) :

        if self.dataEmp :

            jsdata = {}
            cle = ""

            for (numero, 
                 nomnaiss, 
                 nommarit, 
                 prenom, 
                 matrmcdo) in self.dataEmp :


                    cle = numero.replace(' ','')

                    subdir = {'NomNaiss':'', 'NomMarit':'', 'Prenom':'', 'matrMcdo':'', 'absDsn':''}

                    jsdata.setdefault(cle, subdir)

                    jsdata[cle]['NomNaiss'] = nomnaiss
                    jsdata[cle]['NomMarit'] = nommarit
                    jsdata[cle]['Prenom'] = prenom
                    jsdata[cle]['matrMcdo'] = matrmcdo
                    #jsdata[item[0]]['absDsn'] = item[]

        else :
            logging.warning("La requête SQL ne retourne aucune donnée")
        
        logging.debug(jsdata)
        return jsdata


class QueryQPaieLight(object): 

    def __init__(self, chem_base):

        constr = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=' + chem_base
        
         # Requ?te sur la db ListeEmployes
        sSQL_Emp = ("SELECT "
                    "EM.Numero, " 
                    "CL.Texte1 "                                       
                    "FROM Employes EM "
                    "LEFT JOIN CriteresLibres CL " 
                    "ON EM.Numero=CL.NumeroEmploye ")
        logging.debug(sSQL_Emp)
        # Ex?cution requ?te SQL 
        self.dataEmp = []
        try :
            conn = pyodbc.connect(constr, autocommit=True)
            logging.info('Ouverture de {}'.format(chem_base))
            cur = conn.cursor()

            cur.execute(sSQL_Emp)
            self.dataEmp = list(cur)

            # on el?ve les espaces du matricule
            for item in self.dataEmp :
                item[0] = item[0].replace(' ', '')

            conn.close

        except pyodbc.Error :
            print("erreur requete base {} \n {}".format(chem_base,sys.exc_info()[1]))
        except :
            print("erreur ouverture base {} \n {}".format(chem_base,sys.exc_info()[0]))

    def readit(self):
        dic = {}

        # On retourne un dic qui sert de table de correspondance
        # entre les matricules mcdo et les matricules quadra
        # Avec padding dans car l'export des variables SIRE
        # les matricules sont codés sur 6 car.
        # Le matricule mcdo est stocké dans la fiche Employé -> Critères libres
        for x, y in self.dataEmp :
            if y :
                dic[y.zfill(6)] = x
            else :
                dic[x.zfill(6)] = x
        return dic


class QueryDsnEvt (object) :

    def __init__ (self, chem_base, periode) :

        constr = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=' + chem_base

        #per_access = periode.strftime('#%m/%d/%y#')
        
         # Requête sur la table DSN_EVT_ArretReprise
        sSQL = ("SELECT "
                    "NumeroEmploye, " 
                    "MotifArret, "
                    "DateDebut "
                    "FROM DSN_EVT_ArretReprise "
                    "WHERE DateDebut >= {}").format(periode)

        logging.debug(sSQL)
        # Exécution requête SQL 
        self.data = []
        try :
            conn = pyodbc.connect(constr, autocommit=True)
            logging.info('Ouverture de {}'.format(chem_base))
            cur = conn.cursor()

            cur.execute(sSQL)
            self.data = list(cur)

            # on el?ve les espaces du matricule
            for item in self.data :
                item[0] = item[0].replace(' ', '')

            conn.close

        except pyodbc.Error :
            print("erreur requete base {} \n {}".format(chem_base,sys.exc_info()[1]))
        except :
            print("erreur ouverture base {} \n {}".format(chem_base,sys.exc_info()[0]))
    
    def creaDict (self) :

        dic = {}
        for x, y, z in self.data :
            dic.setdefault(x, [])
            dic[x].append(z)
        logging.info("Employes avec ev. DSN : {}".format(len(dic)))
        return dic



        
# abs maladie : 01 - HAB0
# abs AT :      06 - HAB1
# abs pater :   03 - 
# abs mater :   02 - HAB2

# Pour tester la classe QueryQPaieLight
if __name__ == '__main__':

    pp = pprint.PrettyPrinter(indent=4)
    logging.basicConfig(handlers=[logging.basicConfig(level=logging.INFO, format='%(asctime)s -- %(module)s -- %(levelname)s -- %(message)s')])

    periode = datetime(2019, 5, 1)
    QP = QueryPaie()
    QP.connect("assets/thibor.mdb")
    pp.pprint(QP.Employes())
    QP.close()
    # dossier = input("Saisie code dossier : ").zfill(6)
    # db = "//srvquadra/qappli/quadra/database/paie/" + dossier + "/qpaie.mdb"
    # print (db)
    # #perio = datetime(year=2018, month=4, day=1).strftime("#%m/%d/%y#")
    # perio = datetime(year=2018, month=4, day=30)
    # data = QueryPaieSalaries(db, perio)

    # pp.pprint (data.creaDict())

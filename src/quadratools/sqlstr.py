
def req_calc_central(prefix_cli, prefix_frn):
    sql = """
    SELECT
        NBL.CodeJournal, NBL.PeriodeEcriture, NBL.Folio,
        NBL.NbLigneFolio, PRL.ProchaineLigne,
        CLI.DebitClient, CLI.CreditClient,
        FRN.DebitFournisseur, FRN.CreditFournisseur,
        BIL.DebitClasse15, BIL.CreditClasse15,
        EXP.DebitClasse67, EXP.CreditClasse67
    FROM
        (((((SELECT CodeJournal, PeriodeEcriture, Folio, COUNT(*) AS NbLigneFolio
        FROM Ecritures
        WHERE TypeLigne='E'
        GROUP BY CodeJournal, PeriodeEcriture, Folio) NBL
        LEFT JOIN
        (SELECT CodeJournal, PeriodeEcriture, Folio, MAX(LigneFolio) AS ProchaineLigne
        FROM Ecritures
        WHERE TypeLigne='E'
        GROUP BY CodeJournal, PeriodeEcriture, Folio) PRL
        ON NBL.CodeJournal=PRL.CodeJournal AND
            NBL.PeriodeEcriture=PRL.PeriodeEcriture AND
            NBL.Folio=PRL.Folio)
        LEFT JOIN
        (SELECT CodeJournal, PeriodeEcriture, Folio, ROUND(SUM(MontantTenuDebit),2) AS DebitClient, ROUND(SUM(MontantTenuCredit),2) AS CreditClient
        FROM Ecritures
        WHERE TypeLigne='E'
            AND NumeroCompte LIKE '{0}%'
        GROUP BY CodeJournal, PeriodeEcriture, Folio) CLI
        ON NBL.CodeJournal=CLI.CodeJournal AND
            NBL.PeriodeEcriture=CLI.PeriodeEcriture AND
            NBL.Folio=CLI.Folio)
        LEFT JOIN
        (SELECT CodeJournal, PeriodeEcriture, Folio, ROUND(SUM(MontantTenuDebit),2) AS DebitFournisseur, ROUND(SUM(MontantTenuCredit),2) AS CreditFournisseur
        FROM Ecritures
        WHERE TypeLigne='E'
            AND NumeroCompte LIKE '{1}%'
        GROUP BY CodeJournal, PeriodeEcriture, Folio) FRN
        ON NBL.CodeJournal=FRN.CodeJournal AND
            NBL.PeriodeEcriture=FRN.PeriodeEcriture AND
            NBL.Folio=FRN.Folio)
        LEFT JOIN
        (SELECT CodeJournal, PeriodeEcriture, Folio, ROUND(SUM(MontantTenuDebit),2) AS DebitClasse15, ROUND(SUM(MontantTenuCredit),2) AS CreditClasse15
        FROM Ecritures
        WHERE TypeLigne='E'
            AND NumeroCompte LIKE '[1-5]%'
        GROUP BY CodeJournal, PeriodeEcriture, Folio) BIL
        ON NBL.CodeJournal=BIL.CodeJournal AND
            NBL.PeriodeEcriture=BIL.PeriodeEcriture AND
            NBL.Folio=BIL.Folio)
        LEFT JOIN
        (SELECT CodeJournal, PeriodeEcriture, Folio, ROUND(SUM(MontantTenuDebit),2) AS DebitClasse67, ROUND(SUM(MontantTenuCredit),2) AS CreditClasse67
        FROM Ecritures
        WHERE TypeLigne='E'
            AND NumeroCompte LIKE '[6-7]%'
        GROUP BY CodeJournal, PeriodeEcriture, Folio) EXP
        ON NBL.CodeJournal=EXP.CodeJournal AND
            NBL.PeriodeEcriture=EXP.PeriodeEcriture AND
            NBL.Folio=EXP.Folio
    WHERE NBL.CodeJournal<>'AN'
    UNION
    SELECT
        NBL.CodeJournal, #1899-12-30# AS PeriodeEcriture, 0 AS Folio, NBL.NbLigneFolio,
        PRL.ProchaineLigne,
        CLI.DebitClient, CLI.CreditClient,
        FRN.DebitFournisseur, FRN.CreditFournisseur,
        BIL.DebitClasse15, BIL.CreditClasse15,
        0 AS DebitClasse67, 0 AS CreditClasse67
    FROM
        (((((SELECT CodeJournal, COUNT(*) AS NbLigneFolio
        FROM Ecritures
        WHERE TypeLigne='E'
            AND CodeJournal='AN'
        GROUP BY CodeJournal) NBL
        LEFT JOIN
        (SELECT CodeJournal, MAX(LigneFolio) AS ProchaineLigne
        FROM Ecritures
        WHERE TypeLigne='E'
            AND CodeJournal='AN'
        GROUP BY CodeJournal) PRL
        ON NBL.CodeJournal=PRL.CodeJournal)
        LEFT JOIN
        (SELECT CodeJournal, ROUND(SUM(MontantTenuDebit),2) AS DebitClient, ROUND(SUM(MontantTenuCredit),2) AS CreditClient
        FROM Ecritures
        WHERE TypeLigne='E'
            AND CodeJournal='AN'
            AND NumeroCompte LIKE '{0}%'
        GROUP BY CodeJournal) CLI
        ON NBL.CodeJournal=CLI.CodeJournal)
        LEFT JOIN
        (SELECT CodeJournal, ROUND(SUM(MontantTenuDebit),2) AS DebitFournisseur, ROUND(SUM(MontantTenuCredit),2) AS CreditFournisseur
        FROM Ecritures
        WHERE TypeLigne='E'
            AND CodeJournal='AN'
            AND NumeroCompte LIKE '{1}%'
        GROUP BY CodeJournal) FRN
        ON NBL.CodeJournal=FRN.CodeJournal)
        LEFT JOIN
        (SELECT CodeJournal, ROUND(SUM(MontantTenuDebit),2) AS DebitClasse15, ROUND(SUM(MontantTenuCredit),2) AS CreditClasse15
        FROM Ecritures
        WHERE TypeLigne='E'
            AND CodeJournal='AN'
            AND NumeroCompte LIKE '[1-5]%'
        GROUP BY CodeJournal) BIL
        ON NBL.CodeJournal=BIL.CodeJournal)
    """.format(prefix_cli, prefix_frn)
    return sql

# def req_check_central(journal, periode, folio):
#     if journal == "AN":
#         sql = f"""
#             SELECT * FROM Centralisateur 
#             WHERE CodeJournal='{journal}' 
#             """    
#     else:    
#         sql = f"""
#             SELECT * FROM Centralisateur 
#             WHERE CodeJournal='{journal}' 
#             AND Periode=#{periode}# 
#             AND Folio={folio} 
#             """
#     return sql

def req_calc_sld_comptes(dt_exe_fin):
    """
    Requête pour calculer les soldes de tous les comptes
    a partir de la table écriture
    """
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
        WHERE PeriodeEcriture>=#{dt_exe_fin}# 
        AND TypeLigne='E'
        GROUP BY NumeroCompte) N1
        ON NB.NumeroCompte=N1.NumeroCompte
        UNION
        SELECT N.Compte,
            N.debit, N.credit,
            N1.debit, N1.credit,
            N.NbEcritures
        FROM ((
            SELECT
                {} AS Compte, COUNT(*) AS NbEcritures,
                ROUND(SUM(MontantTenuDebit),2) AS debit, ROUND(SUM(MontantTenuCredit),2) as credit
            FROM Ecritures
            WHERE TypeLigne='E'
                AND NumeroCompte LIKE '0*') N
            LEFT JOIN
            (
            SELECT
                "401000" AS Compte, COUNT(*) AS NbEcritures,
                ROUND(SUM(MontantTenuDebit),2) AS debit, ROUND(SUM(MontantTenuCredit),2) as credit
            FROM Ecritures
            WHERE TypeLigne='E'
                AND NumeroCompte LIKE '0*'
                AND PeriodeEcriture>=#2018-12-31#   ) N1
            ON N.Compte=N1.Compte)
        UNION
        SELECT N.Compte,
            N.debit, N.credit,
            N1.debit, N1.credit,
            N.NbEcritures
        FROM ((
            SELECT
                "411000" AS Compte, COUNT(*) AS NbEcritures,
                ROUND(SUM(MontantTenuDebit),2) AS debit, ROUND(SUM(MontantTenuCredit),2) as credit
            FROM Ecritures
            WHERE TypeLigne='E'
                AND NumeroCompte LIKE '9*') N
            LEFT JOIN
            (
            SELECT
                "411000" AS Compte, COUNT(*) AS NbEcritures,
                ROUND(SUM(MontantTenuDebit),2) AS debit, ROUND(SUM(MontantTenuCredit),2) as credit
            FROM Ecritures
            WHERE TypeLigne='E'
                AND NumeroCompte LIKE '9*'
                AND PeriodeEcriture>=#2018-12-31#   ) N1
            ON N.Compte=N1.Compte)        
        """
    return sql

if __name__ == '__main__':
    import logging
#     import pprint
#     from datetime import datetime
#     from quadratools.quadracompta import QueryCompta

#     pp = pprint.PrettyPrinter(indent=4)

    logging.basicConfig(level=logging.DEBUG,
                        format='%(funcName)s\t\t%(levelname)s - %(message)s')

#     # dbpath = "assets/frozen.mdb"
#     dbpath = "assets/ajusted.mdb"
#     # dbpath = "//srvquadra/qappli/quadra/database/cpta/dc/000874/qcompta.mdb"

#     Q = QueryCompta()
#     Q.connect(dbpath)
    sql = req_check_central('AN', "", 1)
    print(sql)

#     Q.close()
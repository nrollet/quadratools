class SqlStr:
    SEL_CENTRAL = """
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
        (SELECT CodeJournal, PeriodeEcriture, Folio, SUM(MontantTenuDebit) AS DebitClient, SUM(MontantTenuCredit) AS CreditClient
        FROM Ecritures
        WHERE TypeLigne='E'
            AND NumeroCompte LIKE '9%'
        GROUP BY CodeJournal, PeriodeEcriture, Folio) CLI
        ON NBL.CodeJournal=CLI.CodeJournal AND
            NBL.PeriodeEcriture=CLI.PeriodeEcriture AND
            NBL.Folio=CLI.Folio)
        LEFT JOIN
        (SELECT CodeJournal, PeriodeEcriture, Folio, SUM(MontantTenuDebit) AS DebitFournisseur, SUM(MontantTenuCredit) AS CreditFournisseur
        FROM Ecritures
        WHERE TypeLigne='E'
            AND NumeroCompte LIKE '0%'
        GROUP BY CodeJournal, PeriodeEcriture, Folio) FRN
        ON NBL.CodeJournal=FRN.CodeJournal AND
            NBL.PeriodeEcriture=FRN.PeriodeEcriture AND
            NBL.Folio=FRN.Folio)
        LEFT JOIN
        (SELECT CodeJournal, PeriodeEcriture, Folio, SUM(MontantTenuDebit) AS DebitClasse15, SUM(MontantTenuCredit) AS CreditClasse15
        FROM Ecritures
        WHERE TypeLigne='E'
            AND NumeroCompte LIKE '[1-5]%'
        GROUP BY CodeJournal, PeriodeEcriture, Folio) BIL
        ON NBL.CodeJournal=BIL.CodeJournal AND
            NBL.PeriodeEcriture=BIL.PeriodeEcriture AND
            NBL.Folio=BIL.Folio)
        LEFT JOIN
        (SELECT CodeJournal, PeriodeEcriture, Folio, SUM(MontantTenuDebit) AS DebitClasse67, SUM(MontantTenuCredit) AS CreditClasse67
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
        NBL.CodeJournal, '' AS PeriodeEcriture, 0 AS Folio, NBL.NbLigneFolio,
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
        (SELECT CodeJournal, SUM(MontantTenuDebit) AS DebitClient, SUM(MontantTenuCredit) AS CreditClient
        FROM Ecritures
        WHERE TypeLigne='E'
            AND CodeJournal='AN'
            AND NumeroCompte LIKE '9%'
        GROUP BY CodeJournal) CLI
        ON NBL.CodeJournal=CLI.CodeJournal)
        LEFT JOIN
        (SELECT CodeJournal, SUM(MontantTenuDebit) AS DebitFournisseur, SUM(MontantTenuCredit) AS CreditFournisseur
        FROM Ecritures
        WHERE TypeLigne='E'
            AND CodeJournal='AN'
            AND NumeroCompte LIKE '0%'
        GROUP BY CodeJournal) FRN
        ON NBL.CodeJournal=FRN.CodeJournal)
        LEFT JOIN
        (SELECT CodeJournal, SUM(MontantTenuDebit) AS DebitClasse15, SUM(MontantTenuCredit) AS CreditClasse15
        FROM Ecritures
        WHERE TypeLigne='E'
            AND CodeJournal='AN'
            AND NumeroCompte LIKE '[1-5]%'
        GROUP BY CodeJournal) BIL
        ON NBL.CodeJournal=BIL.CodeJournal)
    """
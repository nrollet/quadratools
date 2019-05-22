SELECT
    NBL.CodeJournal, NBL.PeriodeEcriture, NBL.Folio,
    NBL.NbLigneFolio, PRL.ProchaineLigne,
    CLI.DebitClient, CLI.CreditClient,
    FRN.DebitFournisseur, FRN.CreditFournisseur,
    BIL.DebitClasse15, BIL.CreditClasse15,
    EXP.DebitClasse67, EXP.CreditClasse67
FROM

/*NbLigneFolio*/
SELECT
    NBL.CodeJournal, 0, NBL.Folio,
    NBL.NbLigneFolio, PRL.ProchaineLigne,
    CLI.DebitClient, CLI.CreditClient,
    FRN.DebitFournisseur, FRN.CreditFournisseur,
    BIL.DebitClasse15, BIL.CreditClasse15,
    EXP.DebitClasse67, EXP.CreditClasse67
FROM
    ((SELECT CodeJournal, Folio, COUNT(*) AS NbLigneFolio
    FROM ECRITURES
    WHERE TypeLigne='E'
        AND CodeJournal='AN'
    GROUP BY CodeJournal) NBL
    LEFT JOIN
    (SELECT CodeJournal, Folio, (MAX(LigneFolio) + 10) AS ProchaineLigne
    FROM Ecritures
    WHERE TypeLigne='E'
    GROUP BY CodeJournal, PeriodeEcriture, Folio) PRL
    ON NBL.CodeJournal=PRL.CodeJournal AND
        NBL.PeriodeEcriture=PRL.PeriodeEcriture AND
        NBL.Folio=PRL.Folio
    )
    LEFT JOIN
    (SELECT CodeJournal, '', 0, SUM(MontantTenuDebit) AS DebitClient, SUM(MontantTenuCredit) AS CreditClient
    FROM ECRITURES
    WHERE TypeLigne='E'
        AND CodeJournal='AN'
        AND NumeroCompte LIKE '9%'
    GROUP BY CodeJournal
)
CLI
/*Soldes Comptes Fournisseurs*/
( 
SELECT CodeJournal, '', 0, SUM(MontantTenuDebit) AS DebitFournisseur, SUM(MontantTenuCredit) AS CreditFournisseur
FROM ECRITURES
WHERE TypeLigne='E'
    AND CodeJournal='AN'
    AND NumeroCompte LIKE '0%'
GROUP BY CodeJournal
)
FRN
/*Soldes Comptes Bilan*/
(SELECT CodeJournal, '', 0, SUM(MontantTenuDebit) AS DebitFournisseur, SUM(MontantTenuCredit) AS CreditFournisseur
FROM ECRITURES
WHERE TypeLigne='E'
    AND CodeJournal='AN'
    AND NumeroCompte LIKE '[1-5]%'
GROUP BY CodeJournal)
BIL
/*Soldes Comptes Exploit*/
(SELECT CodeJournal, '', 0, SUM(MontantTenuDebit) AS DebitFournisseur, SUM(MontantTenuCredit) AS CreditFournisseur
FROM ECRITURES
WHERE TypeLigne='E'
    AND CodeJournal='AN'
    AND NumeroCompte LIKE '[6-7]%'
GROUP BY CodeJournal)
EXP

/*======================*/

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
    GROUP BY CodeJournal)
NBL
    LEFT JOIN
    (SELECT CodeJournal, (MAX(LigneFolio) + 10) AS ProchaineLigne
    FROM Ecritures
    WHERE TypeLigne='E'
        AND CodeJournal='AN'
    GROUP BY CodeJournal)
PRL
    ON NBL.CodeJournal=PRL.CodeJournal
    )
    LEFT JOIN
    (SELECT CodeJournal, SUM(MontantTenuDebit) AS DebitClient, SUM(MontantTenuCredit) AS CreditClient
    FROM Ecritures
    WHERE TypeLigne='E'
        AND CodeJournal='AN'
        AND NumeroCompte LIKE '9*'
    GROUP BY CodeJournal)
CLI
    ON NBL.CodeJournal=CLI.CodeJournal
    )
    LEFT JOIN
    (SELECT CodeJournal, SUM(MontantTenuDebit) AS DebitFournisseur, SUM(MontantTenuCredit) AS CreditFournisseur
    FROM Ecritures
    WHERE TypeLigne='E'
        AND CodeJournal='AN'
        AND NumeroCompte LIKE '0*'
    GROUP BY CodeJournal)
FRN
    ON NBL.CodeJournal=FRN.CodeJournal
    )
    LEFT JOIN
    (SELECT CodeJournal, SUM(MontantTenuDebit) AS DebitClasse15, SUM(MontantTenuCredit) AS CreditClasse15
    FROM Ecritures
    WHERE TypeLigne='E'
        AND CodeJournal='AN'
        AND NumeroCompte LIKE '[1-5]*'
    GROUP BY CodeJournal)
BIL
    ON NBL.CodeJournal=BIL.CodeJournal
    )

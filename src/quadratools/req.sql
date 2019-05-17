SELECT
    NBL.CodeJournal, NBL.PeriodeEcriture, NBL.Folio,
    NBL.NbLigneFolio, PRL.ProchaineLigne,
    CLI.DebitClient, CLI.CreditClient,
    FRN.DebitFournisseur, FRN.CreditFournisseur,
    BIL.DebitClasse15, BIL.CreditClasse15,
    EXP.DebitClasse67, EXP.CreditClasse67
FROM


SELECT CodeJournal, '', 0, COUNT(*) AS NbLigneFolio
FROM ECRITURES
WHERE TypeLigne='E'
    AND CodeJournal='AN'
GROUP BY CodeJournal
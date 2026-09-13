# -*- coding: utf-8 -*-
"""Odczyt paczki raportowej ESEF (.xbri / .zip / .xhtml) - inline XBRL.

To zupelnie inne zrodlo niz scrapowany xlsx: pozycje sa OTAGOWANE nazwami koncepcji
z taksonomii IFRS (ifrs-full:Revenue, ifrs-full:Assets, ...), a nie polskimi etykietami.
Znak, skala i okres sa podane wprost w atrybutach. Dzieki temu mapowanie nie zalezy od
tego, jak spolka nazwala pozycje - jedna tabela koncepcji obsluguje wszystkich emitentow.

    import esef
    fakty = esef.wczytaj("toya-2025-12-31-1-pl.xbri")
    kolumna = esef.na_standard(fakty, "2025-12-31")
"""
import html
import io
import re
import zipfile

CTX = re.compile(r'<xbrli:context id="([^"]+)">(.*?)</xbrli:context>', re.S)
FAKT = re.compile(r'<ix:nonFraction\b([^>]*)>(.*?)</ix:nonFraction>', re.S)


def _atr(s, n):
    m = re.search(n + r'="([^"]*)"', s)
    return m.group(1) if m else None


def _xhtml_z_paczki(sciezka):
    if not zipfile.is_zipfile(sciezka):
        return io.open(sciezka, encoding="utf-8", errors="replace").read()
    with zipfile.ZipFile(sciezka) as z:
        nazwy = [n for n in z.namelist() if n.lower().endswith((".xhtml", ".html"))]
        if not nazwy:
            raise ValueError("W paczce nie ma pliku .xhtml")
        nazwy.sort(key=lambda n: -z.getinfo(n).file_size)      # raport to najwiekszy plik
        return z.read(nazwy[0]).decode("utf-8", "replace")


def _liczba(tresc, format_ixt):
    """Zamienia zawartosc znacznika na liczbe zgodnie z TRANSFORMACJA ixt.

    To jedyne miejsce, gdzie iXBRL naprawde potrafi ugryzc: ta sama tresc "1.234,56"
    znaczy co innego przy 'num-comma-decimal' (przecinek dziesietny) niz przy
    'num-dot-decimal' (kropka dziesietna). Naiwne "usun spacje, przecinek na kropke"
    dziala na polskich raportach i CICHO gubi fakty w angielskich.
    """
    f = (format_ixt or "").split(":")[-1].replace("-", "").lower()
    t = tresc.replace("\xa0", " ").replace("\u2212", "-").strip()
    if f in ("fixedzero", "zerodash", "numdash"):
        return 0.0
    if f == "fixedempty":
        return 0.0
    if t in ("", "-", "\u2013", "\u2014"):
        return 0.0
    if f in ("numcommadecimal", "numspacecomma", "numdotcomma"):
        t = re.sub(r"[ .\u00a0']", "", t).replace(",", ".")
    elif f in ("numdotdecimal", "numspacedot", "numcommadot", "numdotdecimalapos"):
        t = re.sub(r"[ ,\u00a0']", "", t)
    else:
        # brak lub nieznana transformacja - zgadujemy po ostatnim separatorze
        t = t.replace(" ", "")
        if "," in t and "." in t:
            t = (t.replace(".", "").replace(",", ".") if t.rfind(",") > t.rfind(".")
                 else t.replace(",", ""))
        elif "," in t:
            t = t.replace(",", ".")
    return float(t)


def wczytaj(sciezka, tylko_waluta=True):
    """Zwraca (fakty, problemy). 'problemy' to fakty, ktorych NIE udalo sie odczytac -
    lista ma byc pusta; jesli nie jest, nie wolno ufac wynikowi."""
    txt = _xhtml_z_paczki(sciezka)

    jednostki = {}
    for m in re.finditer(r'<xbrli:unit id="([^"]+)">(.*?)</xbrli:unit>', txt, re.S):
        miary = re.findall(r"<xbrli:measure>([^<]+)</xbrli:measure>", m.group(2))
        jednostki[m.group(1)] = {"miary": miary,
                                 "pieniezna": len(miary) == 1 and miary[0].startswith("iso4217:")}

    ctx = {}
    for m in CTX.finditer(txt):
        cid, body = m.group(1), m.group(2)
        inst = re.search(r"<xbrli:instant>([^<]+)", body)
        s_ = re.search(r"<xbrli:startDate>([^<]+)", body)
        e_ = re.search(r"<xbrli:endDate>([^<]+)", body)
        wym = [d[1] for d in re.findall(r'dimension="([^"]+)"[^>]*>([^<]+)<', body)]
        okres = inst.group(1) if inst else (s_.group(1) + ".." + e_.group(1) if s_ and e_ else "?")
        ctx[cid] = {"okres": okres, "koniec": okres.split("..")[-1], "wymiary": wym}

    fakty, problemy = [], []
    for m in FAKT.finditer(txt):
        a, tresc = m.group(1), m.group(2)
        nazwa = _atr(a, "name")
        if "continuedAt" in a:                      # wartosc rozbita na kilka znacznikow
            problemy.append((nazwa, "ix:continuation - nieobslugiwane"))
            continue
        jed = jednostki.get(_atr(a, "unitRef") or "", {"pieniezna": False, "miary": []})
        if tylko_waluta and not jed["pieniezna"]:
            continue                                # np. zysk na akcje (PLN/szt) - standard tego nie uzywa
        czysty = html.unescape(re.sub(r"<[^>]+>", "", tresc))
        try:
            wartosc = _liczba(czysty, _atr(a, "format"))
        except ValueError:
            problemy.append((nazwa, "nie umiem odczytac %r (format=%s)" % (czysty[:30], _atr(a, "format"))))
            continue
        wartosc *= 10 ** int(_atr(a, "scale") or 0)
        if _atr(a, "sign") == "-":
            wartosc = -wartosc
        c = ctx.get(_atr(a, "contextRef"), {"okres": "?", "koniec": "?", "wymiary": []})
        fakty.append({"koncepcja": nazwa, "okres": c["okres"], "koniec": c["koniec"],
                      "wymiary": c["wymiary"], "wartosc": wartosc,
                      "waluta": jed["miary"][0].split(":")[-1] if jed["miary"] else None})
    return fakty, problemy


# --- koncepcja IFRS -> klucz standardu. Wspolna dla WSZYSTKICH emitentow ESEF. ---
# (klucz, znak). Kilka koncepcji moze wpadac do jednego klucza - wtedy sie sumuja.
MAPA = {
    # bilans: aktywa
    "Assets": [("total_assets", 1)],
    "NoncurrentAssets": [("noncurrent_assets", 1)],
    "Goodwill": [("goodwill", 1)],
    "IntangibleAssetsOtherThanGoodwill": [("other_intangible_assets", 1)],
    "PropertyPlantAndEquipment": [("property", 1)],
    "RightofuseAssets": [("property", 1), ("right_to_use_assets", 1)],
    "InvestmentProperty": [("noncurrent_investments", 1)],
    "NoncurrentReceivables": [("noncurrent_receivables", 1)],
    "DeferredTaxAssets": [("other_noncurrent_assets", 1)],
    "OtherNoncurrentAssets": [("other_noncurrent_assets", 1)],
    "CurrentAssets": [("current_assets", 1)],
    "Inventories": [("inventory", 1)],
    "TradeAndOtherCurrentReceivables": [("current_receivables", 1)],
    "CurrentTaxAssetsCurrent": [("current_receivables", 1)],
    "OtherCurrentFinancialAssets": [("current_investments", 1)],
    "CashAndCashEquivalents": [("cash", 1), ("current_investments", 1)],
    "OtherCurrentAssets": [("other_current_assets", 1)],
    "NoncurrentAssetsOrDisposalGroupsClassifiedAsHeldForSale": [("assets_for_sale", 1)],
    # bilans: pasywa
    "EquityAndLiabilities": [("total_equity_liabilities", 1)],
    "Equity": [("capital", 1)],
    "IssuedCapital": [("share_capital", 1)],
    "TreasuryShares": [("own_share", -1)],
    "SharePremium": [("reserve", 1)],
    "RetainedEarnings": [("retained_earnings", 1)],
    "CapitalRedemptionReserve": [("kapitaly_pozostale_nierozdzielone", 1)],
    "ReserveOfExchangeDifferencesOnTranslation": [("kapitaly_pozostale_nierozdzielone", 1)],
    "ReserveOfRemeasurementsOfDefinedBenefitPlans": [("kapitaly_pozostale_nierozdzielone", 1)],
    "OtherReserves": [("kapitaly_pozostale_nierozdzielone", 1)],
    "NoncontrollingInterests": [("nonshare_capital", 1)],
    "NoncurrentLiabilities": [("noncurrent_liabilities", 1)],
    "NoncurrentPayables": [("noncurrent_other_liabilities", 1)],
    "NoncurrentLeaseLiabilities": [("noncurrent_leasing", 1)],
    "DeferredTaxLiabilities": [("noncurrent_other_liabilities", 1)],
    "NoncurrentProvisionsForEmployeeBenefits": [("noncurrent_other_liabilities", 1)],
    "OtherLongtermProvisions": [("noncurrent_other_liabilities", 1)],
    "NoncurrentPortionOfNoncurrentBorrowings": [("noncurrent_borrowings", 1)],
    "CurrentLiabilities": [("current_liabilities", 1)],
    "TradeAndOtherCurrentPayables": [("current_trade_payables", 1)],
    "CurrentLeaseLiabilities": [("current_leasing", 1)],
    "CurrentLoansReceivedAndCurrentPortionOfNoncurrentLoansReceived": [("current_borrowings", 1)],
    "ShorttermBorrowings": [("current_borrowings", 1)],
    "CurrentProvisionsForEmployeeBenefits": [("current_other_liabilities", 1)],
    "CurrentTaxLiabilitiesCurrent": [("current_other_liabilities", 1)],
    "OtherShorttermProvisions": [("current_other_liabilities", 1)],
    "OtherCurrentLiabilities": [("current_other_liabilities", 1)],
    # rachunek wynikow
    "Revenue": [("revenues", 1)],
    "CostOfSales": [("cost_of_sales", 1)],
    "DistributionCosts": [("distribution_expenses", 1)],
    "AdministrativeExpense": [("administrative_expenses", 1)],
    "OtherIncome": [("other_operating_income", 1)],
    "ImpairmentLossImpairmentGainAndReversalOfImpairmentLossDeterminedInAccordanceWithIFRS9":
        [("other_operating_income", -1)],
    "OtherExpenseByFunction": [("other_operating_costs", 1)],
    "ProfitLossFromOperatingActivities": [("ebit", 1)],
    "FinanceIncome": [("finance_income", 1)],
    "FinanceCosts": [("finance_costs", 1)],
    "ShareOfProfitLossOfAssociatesAndJointVenturesAccountedForUsingEquityMethod": [("other_income", 1)],
    "ProfitLossBeforeTax": [("before_tax_profit", 1), ("net_gross_profit", 1)],
    "ProfitLossFromDiscontinuedOperations": [("discontinued_profit", 1)],
    "ProfitLoss": [("net_profit", 1)],
    "ProfitLossAttributableToOwnersOfParent": [("shareholder_net_profit", 1)],
    # przeplywy
    "CashFlowsFromUsedInOperatingActivities": [("operating_cashflow", 1)],
    "AdjustmentsForDepreciationAndAmortisationExpense": [("amortization", 1)],
    "CashFlowsFromUsedInInvestingActivities": [("investing_cashflow", 1)],
    "PurchaseOfPropertyPlantAndEquipmentIntangibleAssetsOtherThanGoodwillInvestmentPropertyAndOtherNoncurrentAssets":
        [("capex", 1)],
    "PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities": [("capex", 1)],
    "PurchaseOfIntangibleAssetsClassifiedAsInvestingActivities": [("capex", 1)],
    "CashFlowsUsedInObtainingControlOfSubsidiariesOrOtherBusinessesClassifiedAsInvestingActivities":
        [("outflows_for_acquisitions", 1)],
    "CashFlowsFromUsedInFinancingActivities": [("financing_cashflow", 1)],
    "ProceedsFromIssuingShares": [("share_capital_cash", 1)],
    "InterestReceivedClassifiedAsInvestingActivities": [("proceed_bank_loans", 1)],
    "InterestPaidClassifiedAsFinancingActivities": [("repaid_bank_loans", 1)],
    "PaymentsOfLeaseLiabilitiesClassifiedAsFinancingActivities": [("lease_liab_payments", 1)],
    "DividendsPaidClassifiedAsFinancingActivities": [("dividend", 1)],
    "IncreaseDecreaseInCashAndCashEquivalentsBeforeEffectOfExchangeRateChanges": [("net_cashflow", 1)],
    # --- dolozone po Neuce: te same pojecia, inne koncepcje taksonomii ---
    "CurrentBorrowingsAndCurrentPortionOfNoncurrentBorrowings": [("current_borrowings", 1)],
    "LongtermBorrowings": [("noncurrent_borrowings", 1)],
    "OtherNoncurrentFinancialAssets": [("noncurrent_investments", 1)],
    "InvestmentAccountedForUsingEquityMethod": [("noncurrent_investments", 1)],
    "OtherNoncurrentFinancialLiabilities": [("noncurrent_other_liabilities", 1)],
    "OtherCurrentFinancialLiabilities": [("current_other_liabilities", 1)],
    "CurrentDerivativeFinancialLiabilities": [("current_other_liabilities", 1)],
    "PurchaseOfInvestmentProperty": [("capex", 1)],
    "OtherGainsLosses": [("other_income", 1)],
    "NoncurrentAssetsOrDisposalGroupsClassifiedAsHeldForSaleOrAsHeldForDistributionToOwners":
        [("assets_for_sale", 1)],
    # "Skup akcji wlasnych" - bierzemy koncepcje z zestawienia zmian w kapitale, bo wystepuje
    # u obu emitentow; odpowiedniki z rachunku przeplywow roznia sie miedzy nimi
    # --- dolozone po Sygnity ---
    "CurrentContractAssets": [("current_receivables", 1)],
    "CurrentPrepaymentsAndCurrentAccruedIncomeIncludingCurrentContractAssets": [("current_receivables", 1)],
    "InvestmentsInSubsidiariesJointVenturesAndAssociates": [("noncurrent_investments", 1)],
    "NoncurrentFinancialAssetsAtAmortisedCost": [("noncurrent_investments", 1)],
    "StatutoryReserve": [("reserve", 1)],
    "NoncurrentProvisions": [("noncurrent_other_liabilities", 1)],
    "CurrentBondsIssuedAndCurrentPortionOfNoncurrentBondsIssued": [("current_obligations", 1)],
    "OtherCurrentNonfinancialLiabilities": [("current_other_liabilities", 1)],
    "CurrentProvisions": [("current_other_liabilities", 1)],
    "CurrentDeferredIncomeIncludingCurrentContractLiabilities": [("current_other_liabilities", 1)],
    "PaymentsToAcquireOrRedeemEntitysShares": [("change_in_bank_loans", 1)],
    # --- dolozone po ATAL ---
    "PropertyPlantAndEquipmentIncludingRightofuseAssets": [("property", 1)],
    "IntangibleAssetsAndGoodwill": [("intangible_assets", 1)],
    "ReserveOfSharebasedPayments": [("kapitaly_pozostale_nierozdzielone", 1)],
    "NoncurrentAdvances": [("noncurrent_other_liabilities", 1)],
    "CurrentAdvances": [("current_other_liabilities", 1)],
    "SellingExpense": [("distribution_expenses", 1)],
    # --- dolozone po CD Projekcie ---
    "OtherIntangibleAssets": [("other_intangible_assets", 1)],
    "InvestmentsInSubsidiaries": [("noncurrent_investments", 1)],
    "NoncurrentPrepaymentsAndNoncurrentAccruedIncomeOtherThanNoncurrentContractAssets": [("other_noncurrent_assets", 1)],
    "OtherNoncurrentReceivables": [("noncurrent_receivables", 1)],
    "CurrentTradeReceivables": [("current_receivables", 1)],
    "OtherCurrentReceivables": [("current_receivables", 1)],
    "CurrentPrepaymentsAndCurrentAccruedIncomeOtherThanCurrentContractAssets": [("other_current_assets", 1)],
    "CurrentHeldtomaturityInvestments": [("current_investments", 1)],
    "RetainedEarningsExcludingProfitLossForReportingPeriod": [("retained_earnings", 1)],
    "RetainedEarningsProfitLossForReportingPeriod": [("year_profit", 1), ("shareholder_net_profit", 1)],
    "OtherNoncurrentPayables": [("noncurrent_other_liabilities", 1)],
    "NoncurrentDeferredIncomeOtherThanNoncurrentContractLiabilities": [("noncurrent_other_liabilities", 1)],
    "TradeAndOtherCurrentPayablesToTradeSuppliers": [("current_trade_payables", 1)],
    "OtherCurrentPayables": [("current_other_liabilities", 1)],
    "CurrentDeferredIncomeOtherThanCurrentContractLiabilities": [("current_other_liabilities", 1)],
    "PaymentsForDevelopmentProjectExpenditure": [("capex", 1)],
    # --- dolozone po Vigo ---
    "CopyrightsPatentsAndOtherIndustrialPropertyRightsServiceAndOperatingRights": [("other_intangible_assets", 1)],
    "IntangibleAssetsUnderDevelopment": [("other_intangible_assets", 1)],
    "NoncurrentAccruedIncomeOtherThanNoncurrentContractAssets": [("other_noncurrent_assets", 1)],
    "OtherReceivables": [("current_receivables", 1)],
    "CurrentReceivablesFromTaxesOtherThanIncomeTax": [("current_receivables", 1)],
    "CurrentAccruedIncomeOtherThanCurrentContractAssets": [("other_current_assets", 1)],
    "NoncurrentPortionOfNoncurrentLoansReceived": [("noncurrent_borrowings", 1)],
    "NoncurrentGovernmentGrants": [("noncurrent_other_liabilities", 1)],
    "CurrentGovernmentGrants": [("current_other_liabilities", 1)],
    "PurchaseOfInterestsInInvestmentsAccountedForUsingEquityMethod": [("outflows_for_acquisitions", 1)],
    # --- dolozone po Selenie ---
    "OtherCurrentNonfinancialAssets": [("other_current_assets", 1)],
    "OtherNoncurrentNonfinancialLiabilities": [("noncurrent_other_liabilities", 1)],
    "ImpairmentLossReversalOfImpairmentLossRecognisedInProfitOrLossLoansAndAdvances": [("other_income", -1)],
    # --- dolozone po Apatorze (wskazane przez same sumy kontrolne) ---
    "NoncurrentPrepayments": [("other_noncurrent_assets", 1)],
    "CurrentPrepayments": [("other_current_assets", 1)],
    "CurrentDerivativeFinancialAssets": [("current_investments", 1)],
    "ReserveOfCashFlowHedges": [("kapitaly_pozostale_nierozdzielone", 1)],
    "CurrentContractLiabilities": [("current_other_liabilities", 1)],
    "CurrentPayablesOnSocialSecurityAndTaxesOtherThanIncomeTax": [("current_other_liabilities", 1)],
    # Apator rozbija amortyzacje na trzy koncepcje zamiast jednej lacznej
    "AdjustmentsForAmortisationExpense": [("amortization", 1)],
    "AdjustmentsForDepreciationExpense": [("amortization", 1)],
}

# Koncepcje ZAPASOWE - uzywane tylko wtedy, gdy klucz nie zostal wypelniony wyzej.
# Emitenci taguja te sama pozycje roznie: Toya konczy przeplywy operacyjne koncepcja
# CashFlowsFromUsedInOperatingActivities, a Sygnity CashFlowsFromUsedInOperations
# (u Toi ta druga oznacza sume POSREDNIA przed podatkiem). Kolejnosc ma znaczenie.
MAPA_ZAPASOWA = {
    # Selena taguje wylacznie przychody z umow z klientami; Sygnity taguje OBIE koncepcje
    # (Revenue = suma), wiec ta druga moze wejsc tylko gdy pierwsza jest nieobecna.
    "RevenueFromContractsWithCustomers": [("revenues", 1)],
    "CashFlowsFromUsedInOperations": [("operating_cashflow", 1)],
    "IncreaseDecreaseInCashAndCashEquivalents": [("net_cashflow", 1)],
    "PurchaseOfTreasuryShares": [("change_in_bank_loans", 1)],
}

# Rozszerzenia emitenta - kazdy tworzy wlasne koncepcje na to, czego nie ma w taksonomii.
# Klucz zewnetrzny to PREFIKS z pliku (toya:, neu:, ...). To jedyna czesc mapowania,
# ktora trzeba robic per spolka - i widac po niej, jak bardzo emitenci sie roznia:
# Toya ma 2 rozszerzenia, Neuca 19.
MAPA_ROZSZERZEN = {
    "toya": {
        "InterestPaidOnBorrowingsClassifiedAsFinancingActivities": [("repaid_bank_loans", 1)],
        "InterestPaidOnLeasingClassifiedAsFinancingActivities": [("repaid_bank_loans", 1)],
    },
    "sel": {
        "ReservedCapital": [("reserve", 1)],
        "RefundLiabilities": [("current_other_liabilities", 1)],   # bonusy obrotowe dla dystrybutorow
        "RevenueFromSaleOfProducts": [],                           # rozbicie przychodow
        "RevenueTradesAndMaterials": [],
    },
    "vgo": {
        "CurrentReceivablesFromTaxesOtherThanIncomeTaxAndShorttermEmployeeBenefitsAccruals":
            [("current_other_liabilities", 1)],
        "Called-upShareCapitalNotPaid": [("current_receivables", 1)],
    },
    "cdprojekt": {
        "NakladyNaPraceRozwojowe": [("intangible_assets", 1)],
        "SupplementaryCapitalFromRetainedEarnings": [("reserve", 1)],
        "InterestsOnBondsReceivedClasifiedAsInvestingActivities": [("proceed_bank_loans", 1)],
        "InterestsOnDepositsReceivedClasifiedAsInvestingActivities": [("proceed_bank_loans", 1)],
        "AdjustmentsForResearchAndDevelopmentExpenseRecognizedAsCostOfSales": [("amortization", 1)],
    },
    "apt": {
        "AdjustmentsForDepreciationExpenseRightOfUseAssets": [("amortization", 1)],
        "ProfitOrLossOnSale": [],                    # rowne wyliczonemu zyskowi ze sprzedazy
    },
    "atal": {
        "OtherReservesAndSupplementaryCapital": [("reserve", 1)],
        "CashInHousingEscrowAccounts": [],          # pozycja "w tym" - srodki na rachunkach powierniczych
    },
    "neu": {
        "NonCurrentInvestmentsInListedEquityInstruments": [("noncurrent_investments", 1)],
        "NoncurrentInvestmentsInNonlistedEquityInstruments": [("noncurrent_investments", 1)],
        "NoncurrentLoans": [("noncurrent_investments", 1)],
        "CurrentLoans": [("current_investments", 1)],
        "NoncurrentContingentConsiderationRecognisedInBusinessCombination": [("noncurrent_other_liabilities", 1)],
        "NoncurrentLiabilitiesArisingFromInsuranceContracts": [("noncurrent_other_liabilities", 1)],
        "CurrentFactoringLiabilities": [("current_other_liabilities", 1)],
        "CurrentContingentConsiderationRecognisedInBusinessCombination": [("current_other_liabilities", 1)],
        "CurrentLiabilitiesArisingFromInsuranceContracts": [("current_other_liabilities", 1)],
        "CurrentAssetsArisingFromInsuranceContracts": [("other_current_assets", 1)],
    },
}


# Nadpisania per emitent dla koncepcji ifrs-full. Ta sama koncepcja potrafi u roznych
# spolek trafiac do INNEGO klucza standardu - to nie jest kwestia taksonomii, tylko tego,
# jak biznesradar poukladal dana spolke. Przyklad: "Pozostale dlugoterminowe zobowiazania
# finansowe" u Neuki ida do noncurrent_other_liabilities, a u ATAL do noncurrent_obligations.
MAPA_NADPISAN = {
    "apt": {
        # u Apatora "pozostale kapitaly rezerwowe" to kapital zapasowy standardu,
        # a nie kapitaly pozostale (u Neuki, Sygnity i CD Projektu jest odwrotnie)
        "OtherReserves": [("reserve", 1)],
    },
    "vgo": {
        # u Vigo pozostale krotkoterminowe aktywa finansowe ida do naleznosci
        "OtherCurrentFinancialAssets": [("current_receivables", 1)],
    },
    "cdprojekt": {
        # CD Projekt trzyma leasing pod ogolnymi "pozostalymi zobowiazaniami finansowymi"
        "OtherNoncurrentFinancialLiabilities": [("noncurrent_leasing", 1)],
        "OtherCurrentFinancialLiabilities": [("current_leasing", 1)],
        # nadwyzka emisyjna nie wchodzi u CDR do kapitalu zapasowego
        "SharePremium": [("kapitaly_pozostale_nierozdzielone", 1)],
        "ImpairmentLossImpairmentGainAndReversalOfImpairmentLossDeterminedInAccordanceWithIFRS9": [],
    },
    "atal": {
        "NoncurrentPayables": [("noncurrent_trade_payables", 1)],
        "OtherNoncurrentFinancialLiabilities": [("noncurrent_obligations", 1)],
        "OtherCurrentFinancialLiabilities": [("current_obligations", 1)],
        # ATAL taguje i SUME rezerw, i jej rozbicie (15 456 + 294 = 15 750).
        # Bierzemy sume, rozbicie odrzucamy - inaczej liczy sie dwa razy.
        "OtherLongtermProvisions": [],
        "NoncurrentProvisionsForEmployeeBenefits": [],
        "OtherShorttermProvisions": [],
        "CurrentProvisionsForEmployeeBenefits": [],
    },
}


def _prefiks_emitenta(fakty):
    """Prefiks wlasnych koncepcji spolki - sluzy za identyfikator emitenta."""
    for f in fakty:
        p = f["koncepcja"].split(":")[0]
        if p not in ("ifrs-full", "esef_cor", "ifrs"):
            return p
    return ""


def na_standard(fakty, koniec_okresu, mnoznik=0.001):
    """Sklada kolumne standardu z faktow o zadanej dacie koncowej.

    mnoznik 0.001, bo ESEF podaje kwoty w ZLOTYCH, a standard w tysiacach.
    Bierzemy wylacznie fakty BEZ wymiarow - wymiarowane to zestawienie zmian
    w kapitale wlasnym, ktore powtarza te same liczby w rozbiciu na skladniki.
    """
    # Ta sama koncepcja bywa otagowana KILKA RAZY w jednym dokumencie: zysk netto
    # pojawia sie w rachunku wynikow, w zestawieniu dochodow calkowitych i jako punkt
    # wyjscia przeplywow. To sa duplikaty tego samego faktu, a nie trzy skladniki -
    # bierzemy je RAZ. Sumujemy tylko RÓŻNE koncepcje wpadajace do jednego klucza.
    emitent = _prefiks_emitenta(fakty)
    out, nieznane, widziane = {}, [], set()
    for f in fakty:
        if f["wymiary"] or f["koniec"] != koniec_okresu:
            continue
        if (f["koncepcja"], f["okres"]) in widziane:
            continue
        widziane.add((f["koncepcja"], f["okres"]))
        prefiks, _, lokalna = f["koncepcja"].partition(":")
        if prefiks == "ifrs-full":
            reguly = MAPA_NADPISAN.get(emitent, {}).get(lokalna, MAPA.get(lokalna))
        else:
            reguly = MAPA_ROZSZERZEN.get(prefiks, {}).get(lokalna)
        if reguly is None:
            nieznane.append(f["koncepcja"])
            continue
        for klucz, znak in reguly:
            out[klucz] = out.get(klucz, 0.0) + f["wartosc"] * znak * mnoznik

    # drugie przejscie: klucze, ktorych nikt nie wypelnil, probujemy z mapy zapasowej
    widziane2 = set()
    for f in fakty:
        if f["wymiary"] or f["koniec"] != koniec_okresu:
            continue
        _, _, lokalna = f["koncepcja"].partition(":")
        reguly = MAPA_ZAPASOWA.get(lokalna)
        if not reguly or (f["koncepcja"], f["okres"]) in widziane2:
            continue
        widziane2.add((f["koncepcja"], f["okres"]))
        for klucz, znak in reguly:
            if klucz not in out:
                out[klucz] = f["wartosc"] * znak * mnoznik
                if f["koncepcja"] in nieznane:
                    nieznane.remove(f["koncepcja"])
    return out, sorted(set(nieznane))


def kolumna_standardu(sciezka, koniec_okresu, bilans_odniesienia=None, mnoznik=0.001):
    """Pelna kolumna standardu z paczki ESEF: mapowanie + reguly + kapital obrotowy.

    Okres odniesienia dla kapitalu obrotowego bierzemy z TEJ SAMEJ paczki - ESEF
    zawiera dane porownawcze, wiec plik jest samowystarczalny.
    """
    fakty, problemy = wczytaj(sciezka)
    biez, nieznane = na_standard(fakty, koniec_okresu, mnoznik)
    rok = int(koniec_okresu[:4])
    odn, _ = na_standard(fakty, "%d-12-31" % (rok - 1), mnoznik)
    # UWAGA: raport ROCZNY przeksztalca dane porownawcze, a standard trzyma bilans
    # PIERWOTNIE zaraportowany. Potwierdzone na Oponeo i CD Projekcie - u CDR naleznosci
    # na 31.12.2024 to 237 390 w raporcie za 2025 i 252 560 w kolumnie IV 2024 standardu.
    # Dlatego przy raporcie rocznym podaj 'bilans_odniesienia' z pliku wzorcowego.
    if bilans_odniesienia:
        odn = {k: (v if isinstance(v, (int, float)) else 0) for k, v in bilans_odniesienia.items()}

    # regula WARUNKOWA - niektorzy emitenci taguja laczna pozycje WNiP wprost
    biez["intangible_assets"] = (biez.get("intangible_assets", 0)
                                 + biez.get("goodwill", 0) + biez.get("other_intangible_assets", 0))
    biez["gross_profit"] = (biez.get("revenues", 0) - biez.get("cost_of_sales", 0)
                            - biez.get("distribution_expenses", 0)
                            - biez.get("administrative_expenses", 0))
    if odn:
        dlug = lambda d: d.get("current_liabilities", 0) - d.get("current_other_liabilities", 0)
        inne = lambda d: d.get("current_other_liabilities", 0) + d.get("reckoning", 0)
        biez["change_in_inventories"] = odn.get("inventory", 0) - biez.get("inventory", 0)
        biez["change_in_receivables"] = odn.get("current_receivables", 0) - biez.get("current_receivables", 0)
        biez["change_in_payables"] = dlug(biez) - dlug(odn)
        biez["change_in_other_assets"] = inne(biez) - inne(odn)
        biez["change_in_working_capital"] = sum(
            biez[k] for k in ("change_in_inventories", "change_in_receivables",
                              "change_in_payables", "change_in_other_assets"))
    biez["balance_date"] = koniec_okresu
    return biez, odn, nieznane, problemy

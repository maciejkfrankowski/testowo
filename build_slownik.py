# -*- coding: utf-8 -*-
"""Buduje slownik.xlsx: Klucze (ze standardu NEU.xlsx) + Aliasy (mapowanie Neuca) + Reguly + Spolki."""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# sciezki wzgledne wzgledem katalogu tego skryptu - dziala po sklonowaniu repo
_TU = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
STD = os.path.join(_TU, "NEU.xlsx")     # dowolny plik w standardzie - zrodlo listy kluczy
OUT = os.path.join(_TU, "slownik.xlsx")

SEKCJE_KLUCZY = {
 **{k: "BILANS_AKTYWA" for k in ["total_assets","noncurrent_assets","intangible_assets","goodwill",
    "other_intangible_assets","property","right_to_use_assets","noncurrent_receivables",
    "noncurrent_investments","other_noncurrent_assets","current_assets","inventory",
    "current_receivables","current_investments","cash","other_current_assets","assets_for_sale"]},
 **{k: "BILANS_PASYWA" for k in ["total_equity_liabilities","capital","share_capital","own_share","reserve",
    "retained_earnings","year_profit","nonshare_capital","noncurrent_liabilities","noncurrent_trade_payables",
    "noncurrent_borrowings","noncurrent_obligations","noncurrent_leasing","noncurrent_other_liabilities",
    "current_liabilities","current_trade_payables","current_borrowings","current_obligations",
    "current_leasing","current_other_liabilities","reckoning"]},
 **{k: "RZIS" for k in ["revenues","cost_of_sales","distribution_expenses","administrative_expenses",
    "gross_profit","other_operating_income","other_operating_costs","ebit","finance_income","finance_costs",
    "other_income","net_gross_profit","extraordinary_profit","before_tax_profit","discontinued_profit",
    "net_profit","shareholder_net_profit"]},
 **{k: "CF" for k in ["operating_cashflow","amortization","change_in_receivables","change_in_inventories",
    "change_in_payables","change_in_other_assets","change_in_working_capital","investing_cashflow","capex",
    "outflows_for_acquisitions","financing_cashflow","share_capital_cash","proceed_bank_loans",
    "repaid_bank_loans","lease_liab_payments","change_in_bank_loans","dividend","net_cashflow"]},
 **{k: "META" for k in ["balance_date","raport_date","share_amount","currency_id"]},
}

# --- Klucze ze standardu ---
wb = openpyxl.load_workbook(STD, read_only=True, data_only=True)
ws = wb["Spr Fin"]
klucze = []
for row in ws.iter_rows(min_row=2, max_row=90, max_col=2, values_only=True):
    k, label = row[0], row[1]
    if not k or str(k).strip() in ("end",):
        continue
    k = str(k).strip()
    klucze.append([k, label, SEKCJE_KLUCZY.get(k, "?")])
wb.close()

# --- Aliasy: spolka, sekcja, nazwa_oryginalna, standard_key, znak, waga, pewnosc, zrodlo, uwagi ---
P, W, D = "PEWNE", "WYSOKA", "DO_WERYFIKACJI"
A = []
ZRODLO = ["NEU Q1 2026"]
def a(sek, nazwa, key, znak=1, pew=P, uw="", spolka="", waga=1):
    A.append([spolka, sek, nazwa, key, znak, waga, pew, ZRODLO[0], uw])

a("*", "Pozycja", "POMIN", 1, P, "wiersz naglowkowy")

# ---------------- RZIS ----------------
a("RZIS","Przychody ze sprzedaży","revenues")
a("RZIS","Koszt własny sprzedaży","cost_of_sales",-1,P,"standard trzyma koszty ze znakiem dodatnim")
a("RZIS","Zysk brutto ze sprzedaży","POMIN",1,P,"PULAPKA: to marza brutto (przychody-KWS). Standardowy gross_profit = po kosztach sprzedazy i zarzadu -> liczony regula")
a("RZIS","Koszty sprzedaży","distribution_expenses",-1)
a("RZIS","Koszty ogólnego zarządu","administrative_expenses",-1)
a("RZIS","Pozostałe przychody operacyjne","other_operating_income")
a("RZIS","Pozostałe koszty operacyjne","other_operating_costs",-1)
a("RZIS","Zmiana stanu odpisów aktualizujących należności i pożyczki","other_operating_income",1,P,"POTWIERDZONE wzorcem: 5024-4157=867 = other_operating_income; odpisy netuja sie z PRZYCHODAMI, nie kosztami")
a("RZIS","Zysk z działalności operacyjnej","ebit")
a("RZIS","Przychody finansowe","finance_income")
a("RZIS","Koszty finansowe","finance_costs",-1)
a("RZIS","Udział w zyskach/stratach jednostek stowarzyszonych","other_income",1,W)
a("RZIS","Pozostałe zyski/straty z inwestycji","other_income",1,W)
a("RZIS","Zysk przed opodatkowaniem","before_tax_profit")
a("RZIS","Zysk przed opodatkowaniem","net_gross_profit",1,P,"standard duplikuje: zysk z dzial. gospodarczej = przed opodatkowaniem")
a("RZIS","Podatek dochodowy","POMIN",1,W,"standard nie ma klucza na podatek - wynika z roznicy before_tax_profit - net_profit")
a("RZIS","Zysk netto","net_profit")
a("RZIS","Zysk przypadający akcjonariuszom jednostki dominującej","shareholder_net_profit",1,P,"NIE zasila year_profit: u spolek MSSF year_profit zostaje pusty, bo caly wynik siedzi w retained_earnings")
a("RZIS","Zysk (strata) przypadający akcjonariuszom niekontrolującym","POMIN",1,P,"zawarte w nonshare_capital w bilansie")
a("RZIS","Średnia ważona liczba akcji zwykłych","share_amount",1,D,"UWAGA: to srednia wazona. Standard oczekuje liczby akcji na dzien bilansowy (tu: kapital podstawowy 4626 tys. / 1 zl = 4 626 000)")
for n in ["Pozycje, które nie mogą być przeniesione do wyniku:",
          "Zysk (strata) z wyceny inwestycji w instrumenty kapitałowe",
          "Pozycje, które nie mogą być przeniesione do wyniku, przed opodatkowaniem",
          "Pozycje, które mogą być przeniesione do wyniku:",
          "Zabezpieczenie przepływów pieniężnych",
          "Różnice kursowe z przeliczenia jednostek działających za granicą",
          "Wycena aktywów finansowych wycenianych przez inne całkowite dochody",
          "Pozycje, które mogą być przeniesione do wyniku, przed opodatkowaniem",
          "Inne całkowite dochody, przed opodatkowaniem",
          "Podatek dochodowy dotyczący składników innych całkowitych dochodów, które mogą być przeniesione do wyniku",
          "Inne całkowite dochody netto","Całkowite dochody ogółem",
          "Całkowite dochody ogółem przypisane akcjonariuszom jednostki dominującej",
          "Całkowite dochody ogółem przypisane akcjonariuszom niekontrolującym",
          "Zysk na 1 akcję podstawowy (w PLN)","Zysk na 1 akcję rozwodniony (w PLN)",
          "Średnia ważona rozwodniona liczba akcji zwykłych"]:
    a("RZIS", n, "POMIN", 1, P, "sekcja innych calkowitych dochodow / EPS - poza zakresem standardu")

# ---------------- CF operacyjny ----------------
a("CF","Zysk przed opodatkowaniem za rok obrotowy","POMIN",1,P,"duplikat z RZiS")
a("CF","Korekty:","POMIN")
a("CF","Zmiany w kapitale obrotowym:","POMIN")
for n in ["Zyski/straty z tytułu różnic kursowych","Odsetki netto",
          "Odpisy aktualizujące aktywa ujęte w wyniku finansowym","Zysk z działalności inwestycyjnej",
          "Koszt programu motywacyjnego","Zmiana stanu rezerw i odpisów aktualizujących zapasy i należności",
          "Udział w stratach jednostek stowarzyszonych","Inne korekty",
          "Korekty z tytułu aktywów netto nabytych jednostek zależnych",
          "Zmiana stanu aktywów ubezpieczeniowych"]:
    a("CF", n, "change_in_other_assets", 1, W, "korekta zbiorcza -> 'Zmiana rezerw i pozostale'")
a("CF","Amortyzacja","amortization")
a("CF","Zmiana stanu zapasów brutto","change_in_inventories")
a("CF","Zmiana stanu należności handlowych oraz pozostałych brutto","change_in_receivables")
a("CF","Zmiana stanu zobowiązań handlowych oraz innych operacyjnych","change_in_payables")
a("CF","Zmiana stanu zobowiązań ubezpieczeniowych","change_in_payables",1,W)
a("CF","Zmiana stanu rozliczeń z tytułu faktoringu należności","change_in_receivables",1,D,"faktoring - decyzja czy do naleznosci czy do dlugu")
a("CF","Zmiana stanu rozliczeń z tytułu faktoringu zobowiązań","change_in_payables",1,D,"j.w.")
a("CF","Zapłacony podatek dochodowy","change_in_other_assets",1,D,"brak klucza na podatek zaplacony w standardzie")
a("CF","Przepływy z działalności operacyjnej przed zmianami w kapitale obrotowym","POMIN",1,P,"suma posrednia")
a("CF","Łącznie korekty zysku oraz zmiany kapitału obrotowego","POMIN",1,P,"suma posrednia")
a("CF","Środki pieniężne wygenerowane w toku działalności operacyjnej","POMIN",1,P,"suma posrednia (przed podatkiem)")
a("CF","Środki pieniężne z działalności operacyjnej (wykorzystane w działalności operacyjnej)","operating_cashflow")

# ---------------- CF inwestycyjny ----------------
a("CF_INW","Wydatki na nabycie wartości niematerialnych oraz rzeczowych aktywów trwałych","capex",-1)
a("CF_INW","Wydatki na nabycie udziałów w jednostkach zależnych (pomniejszone o przejęte środki pieniężne)","outflows_for_acquisitions",-1)
a("CF_INW","Otrzymane odsetki","proceed_bank_loans",1,W,"UWAGA: klucz nazywa sie proceed_bank_loans, ale etykieta w standardzie to 'Odsetki otrzymane'")
a("CF_INW","Środki pieniężne z działalności inwestycyjnej (wykorzystane w działalności inwestycyjnej)","investing_cashflow")
for n in ["Wpływy ze sprzedaży wartości niematerialnych oraz rzeczowych aktywów trwałych",
          "Wpływy ze sprzedaży nieruchomości inwestycyjnych",
          "Wpływy ze sprzedaży instrumentów kapitałowych oraz instrumentów dłużnych",
          "Wydatki na nabycie instrumentów kapitałowych oraz instrumentów dłużnych",
          "Pożyczki udzielone jednostkom niepowiązanym",
          "Otrzymane spłaty pożyczek udzielonych od jednostek niepowiązanych"]:
    a("CF_INW", n, "POMIN", 1, P, "brak odpowiednika - zawarte w investing_cashflow")

# ---------------- CF finansowy ----------------
a("CF_FIN","Wpływy netto z tytułu emisji akcji","share_capital_cash")
a("CF_FIN","Nabycie akcji własnych","change_in_bank_loans",-1,W,"UWAGA: klucz change_in_bank_loans ma w standardzie etykiete 'Skup akcji wlasnych'")
a("CF_FIN","Spłata zobowiązań z tytułu praw do użytkowania","lease_liab_payments",-1)
a("CF_FIN","Odsetki zapłacone","repaid_bank_loans",-1,W,"klucz repaid_bank_loans = etykieta 'Odsetki zaplacone'")
a("CF_FIN","Środki pieniężne z działalności finansowej (wykorzystane w działalności finansowej)","financing_cashflow")
a("CF_FIN","Zmiana netto stanu środków pieniężnych i ich ekwiwalentów","net_cashflow")
for n in ["Wpływy z otrzymanych dotacji","Wpływy z tytułu zaciągnięcia kredytów i pożyczek",
          "Spłaty kredytów i pożyczek",
          "Różnice kursowe otrzymane/(zapłacone) z instrumentów pochodnych ekonomicznie zabezpieczających zobowiązania finansowe, netto",
          "Pozostałe wpływy/wydatki",
          "Zmiana stanu środków pieniężnych i ich ekwiwalentów przed skutkami zmian kursów wymiany",
          "Skutki zmian kursów wymiany, które dotyczą środków pieniężnych i ekwiwalentów środków pieniężnych",
          "Środki pieniężne i ich ekwiwalenty na początek okresu",
          "Środki pieniężne i ich ekwiwalenty na koniec okresu",
          "- o ograniczonej możliwości dysponowania"]:
    a("CF_FIN", n, "POMIN", 1, P, "brak odpowiednika w standardzie")

# ---------------- Bilans: aktywa trwale ----------------
a("AKT_TRW","Rzeczowe aktywa trwałe","property")
a("AKT_TRW","Aktywa z tytułu prawa do użytkowania","property",1,P,"POTWIERDZONE wzorcem: 553250+188778=742028 = property")
a("AKT_TRW","Aktywa z tytułu prawa do użytkowania","right_to_use_assets",1,P,"right_to_use_assets to pozycja 'w tym' - NIE sumuje sie do aktywow trwalych")
a("AKT_TRW","Wartość firmy","goodwill")
a("AKT_TRW","Wartości niematerialne","other_intangible_assets")
a("AKT_TRW","Nieruchomości inwestycyjne","noncurrent_investments",1,D,"alternatywa: property. Standard nie ma osobnego klucza")
a("AKT_TRW","Inwestycje długoterminowe w obligacje","noncurrent_investments")
a("AKT_TRW","Akcje i udziały notowane","noncurrent_investments")
a("AKT_TRW","Należności długoterminowe","noncurrent_receivables")
a("AKT_TRW","Udzielone pożyczki","noncurrent_investments")
a("AKT_TRW","Aktywa z tytułu odroczonego podatku dochodowego","other_noncurrent_assets")
a("AKT_TRW","Aktywa trwałe razem","noncurrent_assets")

# ---------------- Bilans: aktywa obrotowe ----------------
a("AKT_OBR","Zapasy","inventory")
a("AKT_OBR","Należności handlowe oraz pozostałe należności","current_receivables")
a("AKT_OBR","Należności z tytułu bieżącego podatku dochodowego","current_receivables")
a("AKT_OBR","Udzielone pożyczki","current_investments")
a("AKT_OBR","Inne krótkoterminowe aktywa finansowe","current_investments")
a("AKT_OBR","Środki pieniężne i ich ekwiwalenty","cash")
a("AKT_OBR","Środki pieniężne i ich ekwiwalenty","current_investments",1,P,"w standardzie inwestycje krotkoterminowe ZAWIERAJA srodki pieniezne")
a("AKT_OBR","Instrumenty pochodne niewyznaczone do rachunkowości zabezpieczeń","current_investments",1,W)
a("AKT_OBR","Aktywa obrotowe bez aktywów trwałych przeznaczonych do sprzedaży","POMIN",1,P,"suma posrednia")
a("AKT_OBR","Aktywa trwałe zaklasyfikowane jako przeznaczone do sprzedaży","assets_for_sale")
a("AKT_OBR","Aktywa obrotowe razem","current_assets")
a("AKTYWA","Aktywa razem","total_assets")

# ---------------- Bilans: kapital ----------------
a("KAPITAL","Kapitał podstawowy","share_capital")
a("KAPITAL","Akcje własne","own_share",1,P,"juz ze znakiem ujemnym w raporcie - standard tez oczekuje ujemnej")
a("KAPITAL","Kapitał zapasowy ze sprzedaży akcji powyżej ich wartości nominalnej","reserve")
a("KAPITAL","Pozostałe kapitały","kapitaly_pozostale_nierozdzielone",1,P,"POTWIERDZONE wzorcem NEU.xlsx: reserve = 426890 = tylko kapital zapasowy ze sprzedazy akcji")
a("KAPITAL","Niepodzielony wynik finansowy i wynik finansowy roku obrotowego","retained_earnings",1,P,"POTWIERDZONE wzorcem NEU.xlsx: 768518 = CALA pozycja, bez odejmowania wyniku okresu")
a("KAPITAL","Kapitały przypadające akcjonariuszom jednostki dominującej","POMIN",1,P,"suma posrednia")
a("KAPITAL","Kapitały przypadające akcjonariuszom niekontrolującym","nonshare_capital")
a("KAPITAL","Kapitał własny razem","capital")

# ---------------- Bilans: zobowiazania dlugoterminowe ----------------
a("ZOB_DL","Zobowiązania z tytułu prawa do użytkowania","noncurrent_leasing")
a("ZOB_DL","Inne zobowiązania finansowe","noncurrent_other_liabilities",1,P,"POTWIERDZONE wzorcem: noncurrent_borrowings = 0, ta pozycja idzie w inne zobowiazania")
a("ZOB_DL","Zobowiązania z tytułu zakupu udziałów","noncurrent_other_liabilities")
a("ZOB_DL","Rezerwy na zobowiązania z tytułu świadczeń pracowniczych","noncurrent_other_liabilities")
a("ZOB_DL","Inne zobowiązania operacyjne","noncurrent_other_liabilities")
a("ZOB_DL","Długoterminowe zobowiązania ubezpieczeniowe","noncurrent_other_liabilities")
a("ZOB_DL","Zobowiązania długoterminowe razem","noncurrent_liabilities")

# ---------------- Bilans: zobowiazania krotkoterminowe ----------------
a("ZOB_KR","Kredyty i pożyczki","current_borrowings")
a("ZOB_KR","Zobowiązania z tytułu faktoringu należności","current_other_liabilities",1,P,"POTWIERDZONE wzorcem NEU.xlsx I 2026: faktoring NIE wchodzi do current_borrowings")
a("ZOB_KR","Zobowiązania z tytułu prawa do użytkowania","current_leasing")
a("ZOB_KR","Inne zobowiązania finansowe","current_other_liabilities",1,P,"POTWIERDZONE wzorcem: current_borrowings = wylacznie linia Kredyty i pozyczki")
a("ZOB_KR","Zobowiązania z tytułu zakupu udziałów","current_other_liabilities")
a("ZOB_KR","Zobowiązania handlowe oraz inne zobowiązania operacyjne","current_trade_payables")
a("ZOB_KR","Zobowiązania z tytułu bieżącego podatku dochodowego","current_other_liabilities")
a("ZOB_KR","Rezerwy na zobowiązania z tytułu świadczeń pracowniczych","current_other_liabilities")
a("ZOB_KR","Instrumenty pochodne niewyznaczone do rachunkowości zabezpieczeń","current_other_liabilities",1,W)
a("ZOB_KR","Rezerwy na pozostałe zobowiązania","current_other_liabilities")
a("ZOB_KR","Krótkoterminowe zobowiązania ubezpieczeniowe","current_other_liabilities")
a("ZOB_KR","Zobowiązania krótkoterminowe bez zobowiązań wchodzących w skład grup przeznaczonych do sprzedaży","POMIN",1,P,"suma posrednia")
a("ZOB_KR","Zobowiązania wchodzące w skład grup przeznaczonych do sprzedaży","current_other_liabilities")
a("ZOB_KR","Zobowiązania krótkoterminowe razem","current_liabilities")
a("PASYWA","Zobowiązania razem","POMIN",1,P,"suma posrednia")
a("ZOB_KR","Zobowiązania razem","POMIN",1,P,"j.w. - zaleznie od ukladu trafia do ZOB_KR albo PASYWA")
a("PASYWA","Pasywa razem","total_equity_liabilities")


# ================= KGHM =================
ZRODLO[0] = "KGH Q1 2026"
# --- RZiS ---
a("RZIS","Przychody z umów z klientami","revenues")
a("RZIS","Koszty sprzedaży i koszty ogólnego zarządu","administrative_expenses",-1,D,"KGHM nie rozdziela kosztow sprzedazy i zarzadu w raporcie kwartalnym - calosc trafia w koszty zarzadu; rozbicie jest w notach")
a("RZIS","Zysk netto ze sprzedaży","POMIN",1,P,"to jest odpowiednik standardowego gross_profit, ale liczymy go regula z pozycji skladowych")
a("RZIS","Udział w zysku wspólnego przedsięwzięcia wycenianego metodą praw własności","POMIN",1,P,"skladnik sumy 'Wynik z zaangazowania we wspolne przedsiewziecie'")
a("RZIS","Zysk z tytułu odwrócenia utraty wartości pożyczek udzielonych wspólnemu przedsięwzięciu","POMIN",1,P,"skladnik sumy jw.")
a("RZIS","Przychody odsetkowe od pożyczek udzielonych wspólnemu przedsięwzięciu obliczone z zastosowaniem metody efektywnej stopy procentowej","POMIN",1,P,"skladnik sumy jw. - NIE idzie w przychody finansowe")
a("RZIS","Wynik z zaangażowania we wspólne przedsięwzięcie","other_operating_income",1,P,"POTWIERDZONE wzorcem KGH.xlsx: 657+454=1111 = other_operating_income. Caly wynik z JV idzie w przychody OPERACYJNE")
a("RZIS","Pozostałe przychody operacyjne, w tym:","other_operating_income")
a("RZIS","pozostałe odsetki obliczone z zastosowaniem metody efektywnej stopy procentowej","POMIN",1,P,"pozycja 'w tym' - zawarta juz w wierszu nadrzednym, NIE liczyc drugi raz")
a("RZIS","Pozostałe koszty operacyjne, w tym:","other_operating_costs",-1)
a("RZIS","straty z tytułu utraty wartości instrumentów finansowych","POMIN",1,P,"pozycja 'w tym' - zawarta w wierszu nadrzednym")
a("RZIS","Zysk netto przypadający: akcjonariuszom Jednostki Dominującej","shareholder_net_profit",1,P,"j.w. - year_profit zostaje pusty")
a("RZIS","Zysk netto przypadający: na udziały niekontrolujące","POMIN")
a("RZIS","Średnia ważona liczba akcji zwykłych (mln szt.)","share_amount",1,D,"raport podaje w mln szt.; waga 1000 przy mnozniku 1000 daje sztuki. Standard chce liczby akcji na dzien bilansowy",'',1000)
a("RZIS","Zysk na akcję podstawowy i rozwodniony (w PLN)","POMIN")

# --- CF operacyjny ---
a("CF","Zysk przed opodatkowaniem","POMIN",1,P,"duplikat z RZiS - punkt wyjscia rachunku przeplywow")
a("CF","Amortyzacja ujęta w wyniku finansowym","amortization")
for n in ["Udział w zysku wspólnego przedsięwzięcia wycenianego metodą praw własności",
          "Odsetki od pożyczek udzielonych wspólnemu przedsięwzięciu","Pozostałe odsetki",
          "Straty z tytułu utraty wartości rzeczowych aktywów trwałych i wartości niematerialnych",
          "Zysk z tytułu odwrócenia utraty wartości pożyczek udzielonych wspólnemu przedsięwzięciu",
          "Zysk ze zbycia jednostek zależnych","Różnice kursowe, z tego:",
          "Zmiana stanu rezerw na likwidację kopalń, zobowiązań z tytułu programu przyszłych świadczeń pracowniczych oraz pozostałych rezerw",
          "Zmiana stanu pozostałych należności i zobowiązań innych niż kapitał obrotowy",
          "Zmiana stanu aktywów i zobowiązań z tytułu instrumentów pochodnych",
          "Przekwalifikowanie pozostałych całkowitych dochodów do wyniku w związku z realizacją instrumentów pochodnych zabezpieczających",
          "Pozostałe korekty"]:
    a("CF", n, "change_in_other_assets", 1, W, "korekta zbiorcza -> 'Zmiana rezerw i pozostale'")
a("CF","z działalności inwestycyjnej i wyceny środków pieniężnych","POMIN",1,P,"rozwiniecie 'z tego' pozycji Roznice kursowe")
a("CF","z działalności finansowej","POMIN",1,P,"rozwiniecie 'z tego' pozycji Roznice kursowe")
a("CF","Razem wyłączenia przychodów i kosztów","POMIN",1,P,"suma posrednia korekt (zawiera juz amortyzacje)")
a("CF","Podatek dochodowy zapłacony","change_in_other_assets",1,D,"brak klucza na podatek zaplacony w standardzie")
a("CF","Zmiana stanu kapitału obrotowego, w tym:","change_in_wc_nierozdzielony",1,D,"KGHM podaje kapital obrotowy JEDNA kwota - nie da sie rozbic na zapasy/naleznosci/zobowiazania")
a("CF","zmiana stanu zobowiązań handlowych objętych mechanizmami faktoringu odwrotnego","POMIN",1,P,"pozycja 'w tym'")
a("CF","Przepływy pieniężne netto z działalności operacyjnej","operating_cashflow")

# --- CF inwestycyjny ---
a("CF_INW","Wydatki związane z aktywami górniczymi i hutniczymi, w tym:","capex",-1)
a("CF_INW","zapłacone aktywowane odsetki od zadłużenia","POMIN",1,P,"pozycja 'w tym' wewnatrz CAPEX")
a("CF_INW","Wydatki na pozostałe rzeczowe aktywa trwałe i wartości niematerialne","capex",-1)
a("CF_INW","Udzielone zaliczki na rzeczowe aktywa trwałe i wartości niematerialne","POMIN",1,P,"POTWIERDZONE wzorcem KGH.xlsx: capex 1418 = 1147+271, BEZ zaliczek")
a("CF_INW","Odsetki otrzymane z tytułu pożyczek udzielonych wspólnemu przedsięwzięciu","proceed_bank_loans",1,W,"klucz proceed_bank_loans = etykieta 'Odsetki otrzymane'")
a("CF_INW","Przepływy pieniężne netto z działalności inwestycyjnej","investing_cashflow")
for n in ["Wydatki na aktywa finansowe przeznaczone na likwidację kopalń i innych obiektów technologicznych",
          "Wpływy z tytułu spłaty pożyczek udzielonych wspólnemu przedsięwzięciu (kapitał)",
          "Wpływy ze zbycia rzeczowych aktywów trwałych i wartości niematerialnych",
          "Wpływy ze zbycia jednostek zależnych","Pozostałe"]:
    a("CF_INW", n, "POMIN", 1, P, "brak odpowiednika - zawarte w investing_cashflow")

# --- CF finansowy ---
a("CF_FIN","Spłata zobowiązań z tytułu leasingu","lease_liab_payments",-1)
a("CF_FIN","Spłata odsetek, z tego:","repaid_bank_loans",-1,W,"klucz repaid_bank_loans = etykieta 'Odsetki zaplacone'")
a("CF_FIN","od zobowiązań handlowych objętych mechanizmami faktoringu odwrotnego","POMIN",1,P,"rozwiniecie 'z tego'")
a("CF_FIN","z tytułu zadłużenia","POMIN",1,P,"rozwiniecie 'z tego'")
a("CF_FIN","Przepływy pieniężne netto z działalności finansowej","financing_cashflow")
a("CF_FIN","PRZEPŁYWY PIENIĘŻNE NETTO","net_cashflow")
for n in ["Wpływy z tytułu zaciągniętych kredytów i pożyczek","Spłata kredytów i pożyczek","Pozostałe",
          "Różnice kursowe","Stan środków pieniężnych i ich ekwiwalentów na początek okresu",
          "Stan środków pieniężnych i ich ekwiwalentów na koniec okresu, w tym:",
          "środki pieniężne o ograniczonej możliwości dysponowania"]:
    a("CF_FIN", n, "POMIN", 1, W, "brak klucza w standardzie (m.in. zaciagniecie/splata kredytow!)")

# --- Bilans: aktywa trwale ---
a("AKT_TRW","Rzeczowe aktywa trwałe górnicze i hutnicze","property")
a("AKT_TRW","Aktywa niematerialne górnicze i hutnicze","other_intangible_assets")
a("AKT_TRW","Rzeczowe i niematerialne aktywa górnicze i hutnicze","POMIN",1,P,"suma posrednia dwoch pozycji powyzej")
a("AKT_TRW","Pozostałe rzeczowe aktywa trwałe","property")
a("AKT_TRW","Pozostałe aktywa niematerialne","property",1,P,"POTWIERDZONE wzorcem KGH.xlsx: intangible_assets = TYLKO aktywa niematerialne gornicze; pozostale ida w property")
a("AKT_TRW","Pozostałe aktywa rzeczowe i niematerialne","POMIN",1,P,"suma posrednia")
a("AKT_TRW","Wspólne przedsięwzięcie wyceniane metodą praw własności","other_noncurrent_assets",1,P,"POTWIERDZONE wzorcem KGH.xlsx: udzialy wyceniane metoda praw wlasnosci NIE ida w inwestycje dlugoterminowe")
a("AKT_TRW","Pożyczki udzielone wspólnemu przedsięwzięciu","noncurrent_receivables",1,P,"POTWIERDZONE wzorcem KGH.xlsx: pozyczki dla JV = naleznosci dlugoterminowe, nie inwestycje")
a("AKT_TRW","Zaangażowanie we wspólne przedsięwzięcie","POMIN",1,P,"suma posrednia")
a("AKT_TRW","Pochodne instrumenty finansowe","noncurrent_investments")
a("AKT_TRW","Inne instrumenty finansowe wyceniane w wartości godziwej","noncurrent_investments")
a("AKT_TRW","Inne instrumenty finansowe wyceniane w zamortyzowanym koszcie","noncurrent_investments")
a("AKT_TRW","Instrumenty finansowe razem","POMIN",1,P,"suma posrednia")
a("AKT_TRW","Pozostałe aktywa niefinansowe","other_noncurrent_assets")
a("AKT_TRW","Aktywa trwałe","noncurrent_assets",1,P,"u KGHM to SUMA zamykajaca blok, nie naglowek")

# --- Bilans: aktywa obrotowe ---
a("AKT_OBR","Należności od odbiorców, w tym:","current_receivables")
a("AKT_OBR","należności od odbiorców wyceniane w wartości godziwej przez wynik finansowy","POMIN",1,P,"pozycja 'w tym'")
a("AKT_OBR","Należności z tytułu podatków","current_receivables")
a("AKT_OBR","Pochodne instrumenty finansowe","current_investments")
a("AKT_OBR","Pozostałe aktywa finansowe","current_investments")
a("AKT_OBR","Pozostałe aktywa niefinansowe","other_current_assets")
a("AKT_OBR","Aktywa obrotowe","current_assets",1,P,"suma zamykajaca blok")
a("AKTYWA","RAZEM AKTYWA","total_assets")

# --- Bilans: kapital ---
a("KAPITAL","Kapitał akcyjny","share_capital")
a("KAPITAL","Kapitał z tytułu wyceny instrumentów finansowych","kapitaly_pozostale_nierozdzielone",1,P,"POTWIERDZONE wzorcem KGH.xlsx: reserve = 0. Kapitaly z wyceny/OCI NIE sa ujmowane w rozbiciu kapitalu wlasnego")
a("KAPITAL","Zakumulowane pozostałe całkowite dochody inne niż z tytułu wyceny instrumentów finansowych","kapitaly_pozostale_nierozdzielone",1,P,"j.w. - pozycja OCI poza rozbiciem")
a("KAPITAL","Zyski zatrzymane","retained_earnings",1,P,"POTWIERDZONE wzorcem KGH.xlsx: CALA pozycja idzie w retained_earnings, year_profit zostaje pusty")
a("KAPITAL","Kapitał własny akcjonariuszy Jednostki Dominującej","POMIN",1,P,"suma posrednia")
a("KAPITAL","Kapitał własny udziałowców niekontrolujących","nonshare_capital")
a("KAPITAL","Kapitał własny","capital",1,P,"suma zamykajaca blok")

# --- Bilans: zobowiazania dlugoterminowe ---
a("ZOB_DL","Zobowiązania z tytułu kredytów, pożyczek oraz leasingu","noncurrent_borrowings",1,D,"KGHM laczy kredyty z leasingiem; standard ma osobny klucz noncurrent_leasing")
a("ZOB_DL","Zobowiązania z tytułu dłużnych papierów wartościowych","noncurrent_obligations")
a("ZOB_DL","Pochodne instrumenty finansowe","noncurrent_other_liabilities")
a("ZOB_DL","Zobowiązania z tytułu świadczeń pracowniczych","noncurrent_other_liabilities")
a("ZOB_DL","Rezerwy na koszty likwidacji kopalń i innych obiektów technologicznych","noncurrent_other_liabilities")
a("ZOB_DL","Zobowiązania z tytułu odroczonego podatku dochodowego","noncurrent_other_liabilities")
a("ZOB_DL","Pozostałe zobowiązania","noncurrent_other_liabilities")
a("ZOB_DL","Zobowiązania długoterminowe","noncurrent_liabilities",1,P,"suma zamykajaca blok")

# --- Bilans: zobowiazania krotkoterminowe ---
a("ZOB_KR","Zobowiązania z tytułu kredytów, pożyczek oraz leasingu","current_borrowings",1,D,"j.w. - kredyty razem z leasingiem")
a("ZOB_KR","Zobowiązania z tytułu dłużnych papierów wartościowych","current_obligations")
a("ZOB_KR","Pochodne instrumenty finansowe","current_other_liabilities")
a("ZOB_KR","Zobowiązania wobec dostawców i pozostałe","current_trade_payables")
a("ZOB_KR","Zobowiązania z tytułu świadczeń pracowniczych","current_other_liabilities")
a("ZOB_KR","Zobowiązania z tytułu podatków","current_other_liabilities")
a("ZOB_KR","Rezerwy na zobowiązania i inne obciążenia","current_other_liabilities")
a("ZOB_KR","Pozostałe zobowiązania","current_other_liabilities")
a("ZOB_KR","Zobowiązania krótkoterminowe","current_liabilities",1,P,"suma zamykajaca blok")
a("PASYWA","Zobowiązanie długo i krótkoterminowe","POMIN",1,P,"suma posrednia")
a("PASYWA","RAZEM ZOBOWIĄZANIA I KAPITAŁ WŁASNY","total_equity_liabilities")


# ================= AOL - uklad ustawy o rachunkowosci (UoR) =================
# Aliasy nazwane wg wzorca ustawowego - powinny dzialac dla kazdej kolejnej spolki na UoR.
ZRODLO[0] = "AOL Q1 2026"
a("*","⚠️ UWAGA: koszty w tym sprawozdaniu mogą być podane jako liczby DODATNIE (bez znaku minus) — sprawdź konwencję znaków przed sumowaniem/analizą tych danych.","POMIN",1,P,"komentarz w pliku, nie pozycja sprawozdania")

# --- RZiS, wariant POROWNAWCZY ---
a("RZIS","A. Przychody netto ze sprzedaży i zrównane z nimi","revenues")
a("RZIS","- w tym od jednostek powiązanych nieobjętych metodą konsolidacji pełnej","POMIN",1,P,"pozycja 'w tym'")
a("RZIS","I. Przychody netto ze sprzedaży produktów","POMIN",1,P,"rozwiniecie pozycji A")
a("RZIS","B. Koszty działalności operacyjnej","cost_of_sales",1,P,"POTWIERDZONE wzorcem AOL.xlsx: w wariancie porownawczym CALOSC kosztow operacyjnych idzie w cost_of_sales, a koszty sprzedazy i zarzadu zostaja PUSTE")
for n in ["I. Amortyzacja","II. Zużycie materiałów i energii","III. Usługi obce","IV. Podatki i opłaty, w tym:",
          "- podatek akcyzowy","V. Wynagrodzenia","VI. Ubezpieczenia społeczne i inne świadczenia, w tym:",
          "- emerytalne","VII. Pozostałe koszty rodzajowe"]:
    a("RZIS", n, "POMIN", 1, P, "uklad rodzajowy - rozwiniecie pozycji B, NIE liczyc osobno")
a("RZIS","C. Zysk (strata) ze sprzedaży (A-B)","POMIN",1,P,"gross_profit liczony regula z zaokraglonych skladnikow (3458-2661=797, a nie 796)")
a("RZIS","D. Pozostałe przychody operacyjne","other_operating_income")
a("RZIS","I. Zysk z tytułu rozchodu niefinansowych aktywów trwałych","POMIN",1,P,"rozwiniecie D")
a("RZIS","II. Inne przychody operacyjne","POMIN",1,P,"rozwiniecie D")
a("RZIS","E. Pozostałe koszty operacyjne","other_operating_costs")
a("RZIS","I. Strata ze zbycia niefinansowych aktywów trwałych","POMIN",1,P,"rozwiniecie E")
a("RZIS","II. Inne koszty operacyjne","POMIN",1,P,"rozwiniecie E")
a("RZIS","F. Zysk (strata) z działalności operacyjnej (C+D-E)","ebit")
a("RZIS","G. Przychody finansowe","finance_income")
a("RZIS","H. Koszty finansowe","finance_costs")
for n in ["I. Odsetki, w tym:","- od jednostek powiązanych","- dla jednostek powiązanych",
          "II. Zysk ze zbycia inwestycji","III. Aktualizacja wartości aktywów finansowych","IV. Inne",
          "II. Strata z tytułu rozchodu aktywów finansowych, w tym:","- w jednostkach powiązanych"]:
    a("RZIS", n, "POMIN", 1, P, "rozwiniecie przychodow/kosztow finansowych")
a("RZIS","I. Zysk (strata) z działalności gospodarczej (F+G-H)","net_gross_profit")
a("RZIS","J. Odpis wartości firmy","POMIN",1,W,"brak klucza w standardzie")
a("RZIS","K. Zysk (strata) brutto (I-J)","before_tax_profit")
a("RZIS","L. Podatek dochodowy","POMIN",1,P,"standard nie ma klucza na podatek")
a("RZIS","M. Pozostałe obowiązkowe zmniejszenia zysku (zwiększenia straty)","POMIN")
a("RZIS","N. Zysk (strata) netto (K-L-M)","net_profit")
a("RZIS","N. Zysk (strata) netto (K-L-M)","shareholder_net_profit",1,P,"brak udzialow niekontrolujacych w jednostkowym UoR")

# --- Przeplywy (metoda posrednia UoR) ---
a("CF","A. Przepływy środków pieniężnych z działalności operacyjnej","POMIN",1,P,"powtorzenie pozycji A.III")
a("CF","I. Zysk (strata) netto","POMIN",1,P,"punkt wyjscia - juz w RZiS")
a("CF","II. Korekty razem","POMIN",1,P,"suma posrednia")
a("CF","1. Amortyzacja","amortization")
a("CF","6. Zmiana stanu zapasów","change_in_inventories")
a("CF","7. Zmiana stanu należności","change_in_receivables")
a("CF","8. Zmiana stanu zobowiązań krótkoterminowych, z wyjątkiem pożyczek i kredytów","change_in_payables")
for n in ["2. Zyski (straty) z tytułu różnic kursowych","3. Odsetki i udziały w zyskach (dywidendy)",
          "4. Zysk (strata) z działalności inwestycyjnej","5. Zmiana stanu rezerw",
          "9. Zmiana stanu rozliczeń międzyokresowych","10. Inne korekty z działalności operacyjnej"]:
    a("CF", n, "change_in_other_assets", 1, W, "korekta zbiorcza")
a("CF","III. Przepływy pieniężne netto z działal. operacyjnej (I±II)","operating_cashflow")
a("CF_FIN","D. Przepływy pieniężne netto razem (A.III±B.III±C.III)","net_cashflow")
for n in ["E. Bilansowa zmiana stanu środków pieniężnych, w tym","- zmiana stanu środków pieniężnych z tyt. różnic kurs.",
          "F. Środki pieniężne na początek okresu","G. Środki pieniężne na koniec okresu (F±D), w tym:"]:
    a("CF_FIN", n, "POMIN", 1, P, "stany srodkow pienieznych - poza zakresem standardu")

a("CF_INW","B. Przepływy środków pieniężnych z działal. inwest.","POMIN",1,P,"powtorzenie B.III")
a("CF_INW","1. Nabycie wartości niematerialnych i prawnych oraz rzeczowych aktywów trwałych","capex")
a("CF_INW","- odsetki","proceed_bank_loans",1,W,"klucz proceed_bank_loans = etykieta 'Odsetki otrzymane'")
a("CF_INW","III. Przepływy pieniężne netto z działalności inwestycyjnej (I-II)","investing_cashflow")
for n in ["I. Wpływy","II. Wydatki","1. Zbycie wartości niematerialnych i prawnych oraz rzeczowych aktywów trwałych",
          "2. Z aktywów finansowych, w tym:","b) w pozostałych jednostkach"]:
    a("CF_INW", n, "POMIN", 1, P, "rozwiniecie przeplywow inwestycyjnych")

a("CF_FIN","C. Przepływy środków pieniężnych z działalności finansowej","POMIN",1,P,"powtorzenie C.III")
a("CF_FIN","1. Wpływy netto z wydania udziałów (emisji akcji) i innych instrum. kapit. oraz dopłat do kapitału","share_capital_cash")
a("CF_FIN","1. Dywidendy i inne wypłaty na rzecz właścicieli","dividend")
a("CF_FIN","III. Przepływy pieniężne netto z działalności finansowej (I-II)","financing_cashflow")
for n in ["I. Wpływy","II. Wydatki","2. Inne wpływy finansowe","2. Z tytułu innych zobowiązań finansowych","3. Inne wydatki finansowe"]:
    a("CF_FIN", n, "POMIN", 1, P, "rozwiniecie przeplywow finansowych")

# --- Bilans: aktywa trwale ---
a("AKT_TRW","A. Aktywa trwałe","noncurrent_assets")
a("AKT_TRW","I. Wartości niematerialne i prawne","other_intangible_assets")
a("AKT_TRW","II. Rzeczowe aktywa trwałe","property")
a("AKT_TRW","III. Należności długoterminowe","noncurrent_receivables")
a("AKT_TRW","IV. Inwestycje długoterminowe","noncurrent_investments")
a("AKT_TRW","V. Długoterminowe rozliczenia międzyokresowe","other_noncurrent_assets")
for n in ["1. Koszty zakończonych prac rozwojowych","2. Inne wartości niematerialne i prawne","1. Środki trwałe",
          "a) budynki, lokale i obiekty inżynierii lądowej i wodnej","b) urządzenia techniczne i maszyny",
          "c) środki transportu","d) inne środki trwałe","1. Od pozostałych jednostek",
          "1. Długoterminowe aktywa finansowe","a) w jednostkach powiązanych","- udziały i akcje",
          "1. Aktywa z tytułu odroczonego podatku dochodowego","2. Inne rozliczenia międzyokresowe"]:
    a("AKT_TRW", n, "POMIN", 1, P, "rozwiniecie pozycji nadrzednej")

# --- Bilans: aktywa obrotowe ---
a("AKT_OBR","B. Aktywa obrotowe","current_assets")
a("AKT_OBR","I. Zapasy","inventory")
a("AKT_OBR","II. Należności krótkoterminowe","current_receivables")
a("AKT_OBR","III. Inwestycje krótkoterminowe","current_investments")
a("AKT_OBR","b) środki pieniężne i inne aktywa pieniężne","cash",1,P,"w standardzie inwestycje krotkoterminowe ZAWIERAJA srodki pieniezne")
a("AKT_OBR","IV. Krótkoterminowe rozliczenia międzyokresowe","other_current_assets")
for n in ["1. Zaliczki na dostawy i usługi","1. Należności od pozostałych jednostek",
          "a) z tytułu dostaw i usług, o okresie spłaty:","- do 12 miesięcy","- powyżej 12 miesięcy",
          "b) z tytułu podatków, dotacji, ceł, ubezpieczeń społecznych i zdrowotnych oraz innych świadczeń",
          "c) inne","1. Krótkoterminowe aktywa finansowe","a) w pozostałych jednostkach",
          "- inne papiery wartościowe","- udzielone pożyczki","- środki pieniężne w kasie i na rachunkach",
          "- inne środki pieniężne","C. Należne wpłaty na kapitał (fundusz) podstawowy","D. Udziały (akcje) własne"]:
    a("AKT_OBR", n, "POMIN", 1, P, "rozwiniecie pozycji nadrzednej")

# --- Bilans: kapital wlasny ---
a("KAPITAL","A. Kapitał własny","capital")
a("KAPITAL","I. Kapitał podstawowy","share_capital")
a("KAPITAL","II. Kapitał zapasowy","reserve")
a("KAPITAL","III. Pozostałe kapitały rezerwowe","kapitaly_pozostale_nierozdzielone",1,P,"POTWIERDZONE wzorcem: kapitaly rezerwowe poza rozbiciem, jak OCI u spolek MSSF")
a("KAPITAL","IV. Zysk (strata) z lat ubiegłych","retained_earnings")
a("KAPITAL","V. Zysk (strata) netto","year_profit",1,P,"UoR rozdziela wynik okresu od lat ubieglych, wiec year_profit JEST wypelniany (inaczej niz u MSSF)")

# --- Bilans: rezerwy i zobowiazania ---
a("PASYWA","B. Zobowiązania i rezerwy na zobowiązania","POMIN",1,P,"suma posrednia")
a("PASYWA","I. Rezerwy na zobowiązania","reckoning",1,P,"POTWIERDZONE wzorcem AOL.xlsx: reckoning = rezerwy + rozliczenia miedzyokresowe (337+152=490)")
for n in ["1. Rezerwa z tytułu odroczonego podatku dochodowego","2. Rezerwa na świadczenia emerytalne i podobne",
          "- krótkoterminowa","- długoterminowa","3. Pozostałe rezerwy","- krótkoterminowe"]:
    a("PASYWA", n, "POMIN", 1, P, "rozwiniecie rezerw")

a("ZOB_DL","II. Zobowiązania długoterminowe","noncurrent_liabilities")
for n in ["1. Wobec pozostałych jednostek","a) kredyty i pożyczki","b) inne zobowiązania finansowe"]:
    a("ZOB_DL", n, "POMIN", 1, P, "rozwiniecie zobowiazan dlugoterminowych")

a("ZOB_KR","III. Zobowiązania krótkoterminowe","current_liabilities")
a("ZOB_KR","a) kredyty i pożyczki","current_borrowings")
a("ZOB_KR","c) z tytułu dostaw i usług, o okresie wymagalności:","current_trade_payables")
a("ZOB_KR","d) z tytułu podatków, ceł, ubezpieczeń i innych świadczeń","current_other_liabilities")
a("ZOB_KR","e) z tytułu wynagrodzeń","current_other_liabilities")
a("ZOB_KR","f) inne","current_other_liabilities")
a("ZOB_KR","IV. Rozliczenia międzyokresowe","reckoning")
a("ZOB_KR","b) inne zobowiązania finansowe","current_other_liabilities",1,W,"jak u Neuki: inne zobowiazania finansowe nie sa dlugiem")
for n in ["1. Wobec pozostałych jednostek","- do 12 miesięcy","1. Inne rozliczenia międzyokresowe","- krótkoterminowe"]:
    a("ZOB_KR", n, "POMIN", 1, P, "rozwiniecie zobowiazan krotkoterminowych")

# --- Reguly obliczane ---
REGULY = [
 [1,"intangible_assets","goodwill + other_intangible_assets","NIE","Standard: WNiP ogolem = wartosc firmy + pozostale WNiP"],
 [2,"gross_profit","revenues - cost_of_sales - distribution_expenses - administrative_expenses","NIE",
    "UWAGA: standardowy 'Zysk ze sprzedazy' to NIE jest 'Zysk brutto ze sprzedazy' z raportu spolki"],
 [3,"ebit","gross_profit + other_operating_income - other_operating_costs","TAK",
    "Regula WARUNKOWA: liczona tylko gdy spolka nie pokazuje linii EBIT (KGHM idzie od zysku ze sprzedazy prosto do zysku przed opodatkowaniem). Gdy spolka podaje EBIT wprost (Neuca), zostaje wartosc raportowana"],
 [4,"change_in_working_capital","change_in_receivables + change_in_inventories + change_in_payables + change_in_other_assets","NIE",
    "UWAGA: przy fladze kapital_obrotowy_z_bilansu=TAK te 4 klucze sa i tak nadpisywane wyliczeniem z bilansu"],
]

# --- Konfiguracja spolek ---
SEKCJE_CFG = ("Income Statement (Rachunek Zyskow i Strat / P&L / Statement of Comprehensive Income)=>RZIS"
 "|Cash Flow Statement (Rachunek Przeplywow Pienieznych)=>CF"
 "|Przepływy środków pieniężnych z działalności inwestycyjnej=>CF_INW"
 "|Przepływy środków pieniężnych z działalności finansowej=>CF_FIN"
 "|Balance Sheet (Bilans / Statement of Financial Position)=>BILANS"
 "|Aktywa trwałe=>AKT_TRW"
 "|$Aktywa trwałe razem=>AKT_TRW"
 "|Aktywa obrotowe=>AKT_OBR"
 "|$Aktywa obrotowe razem=>AKT_OBR"
 "|$Aktywa razem=>AKTYWA"
 "|Kapitał własny=>KAPITAL"
 "|$Kapitał własny razem=>KAPITAL"
 "|Zobowiązania długoterminowe=>ZOB_DL"
 "|$Zobowiązania długoterminowe razem=>ZOB_DL"
 "|Zobowiązania krótkoterminowe=>ZOB_KR"
 "|$Zobowiązania krótkoterminowe razem=>ZOB_KR"
 "|$Pasywa razem=>PASYWA")
SEKCJE_KGHM = ("Income Statement (Rachunek Zyskow i Strat / P&L / Statement of Comprehensive Income)=>RZIS"
 "|Cash Flow Statement (Rachunek Przeplywow Pienieznych)=>CF"
 "|Przepływy pieniężne z działalności operacyjnej=>CF"
 "|Przepływy pieniężne z działalności inwestycyjnej=>CF_INW"
 "|Przepływy pieniężne z działalności finansowej=>CF_FIN"
 "|Balance Sheet (Bilans / Statement of Financial Position)=>BILANS"
 "|AKTYWA=>AKTYWA"
 "|$Aktywa trwałe=>AKT_TRW"
 "|$Aktywa obrotowe=>AKT_OBR"
 "|$RAZEM AKTYWA=>AKTYWA"
 "|ZOBOWIĄZANIA I KAPITAŁ WŁASNY=>PASYWA"
 "|$Kapitał własny=>KAPITAL"
 "|$Zobowiązania długoterminowe=>ZOB_DL"
 "|$Zobowiązania krótkoterminowe=>ZOB_KR"
 "|$RAZEM ZOBOWIĄZANIA I KAPITAŁ WŁASNY=>PASYWA")
SEKCJE_AOL = ("Income Statement (Rachunek Zyskow i Strat / P&L / Statement of Comprehensive Income)=>RZIS"
 "|Cash Flow Statement (Rachunek Przeplywow Pienieznych)=>CF"
 "|A. Przepływy środków pieniężnych z działalności operacyjnej=>CF"
 "|III. Przepływy pieniężne netto z działal. operacyjnej (I±II)=>CF"
 "|B. Przepływy środków pieniężnych z działal. inwest.=>CF_INW"
 "|III. Przepływy pieniężne netto z działalności inwestycyjnej (I-II)=>CF_INW"
 "|C. Przepływy środków pieniężnych z działalności finansowej=>CF_FIN"
 "|III. Przepływy pieniężne netto z działalności finansowej (I-II)=>CF_FIN"
 "|Balance Sheet (Bilans / Statement of Financial Position)=>BILANS"
 "|A. Aktywa trwałe=>AKT_TRW"
 "|B. Aktywa obrotowe=>AKT_OBR"
 "|$Aktywa razem=>AKTYWA"
 "|A. Kapitał własny=>KAPITAL"
 "|B. Zobowiązania i rezerwy na zobowiązania=>PASYWA"
 "|II. Zobowiązania długoterminowe=>ZOB_DL"
 "|III. Zobowiązania krótkoterminowe=>ZOB_KR"
 "|$Pasywa razem=>PASYWA")
SPOLKI = [["NEU","Neuca","Neuca za I kwartał 2026 roku_sprawozdania.xlsx","Sprawozdania",1,"2,3",1,
           "I kw 2026 / 31.03.2026|I kw 2025 / 31.12.2025", SEKCJE_CFG,
           "TAK","dane w tys. PLN. Uklad: naglowki otwieraja bloki bilansu"],
          ["AOL","AOL","788_AO_IQ_2026_2026-05-14_final_sprawozdania.xlsx","Sprawozdania",1,"2,3",0.001,"I kw 2026 / 31.03.2026|I kw 2025 / 31.03.2025", SEKCJE_AOL, "NIE","uklad USTAWY O RACHUNKOWOSCI. Dane w ZLOTYCH -> mnoznik 0.001. UWAGA: kolumna porownawcza to 31.03.2025 (rok wczesniej), wiec kapitalu obrotowego NIE da sie policzyc z bilansu - flaga NIE, uzywane sa pozycje z rachunku przeplywow"],
          ["KGH","KGHM","skonsolidowany raport kwartalny KGHM 1_2026_sprawozdania.xlsx","Sprawozdania",1,"2,3",1000,
           "I kw 2026 / 31.03.2026|I kw 2025 / 31.12.2025", SEKCJE_KGHM,
           "TAK","dane w mln PLN -> mnoznik 1000. Uklad: sumy ZAMYKAJA bloki bilansu, stad znaczniki $"]]

# ================= zapis =================
HDR_FILL = PatternFill("solid", fgColor="1F3864"); HDR_FONT = Font(bold=True, color="FFFFFF")
WARN = PatternFill("solid", fgColor="FFEB9C"); GREY = PatternFill("solid", fgColor="F2F2F2")

wb = openpyxl.Workbook()
def sheet(name, header, rows, widths):
    ws = wb.create_sheet(name) if wb.sheetnames != ["Sheet"] or name != "Klucze" else wb.active
    ws.title = name
    ws.append(header)
    for r in rows: ws.append(r)
    for c in range(1, len(header) + 1):
        cell = ws.cell(row=1, column=c); cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    return ws

sheet("Klucze", ["standard_key","etykieta_pl","sekcja"], klucze, [32,52,18])
ws = sheet("Aliasy", ["spolka","sekcja","nazwa_oryginalna","standard_key","znak","waga","pewnosc","zrodlo","uwagi"],
           A, [10,12,72,30,7,7,16,14,70])
for r in range(2, ws.max_row + 1):
    pew = ws.cell(row=r, column=7).value
    sk = ws.cell(row=r, column=4).value
    if pew == "DO_WERYFIKACJI":
        for c in range(1, 10): ws.cell(row=r, column=c).fill = WARN
    elif sk == "POMIN":
        for c in range(1, 10): ws.cell(row=r, column=c).fill = GREY
sheet("Reguly_obliczane", ["kolejnosc","standard_key","formula","tylko_gdy_brak","uwagi"], REGULY, [11,30,72,16,80])
sheet("Spolki", ["spolka","nazwa","plik","arkusz","kol_nazwa","kol_wartosci","mnoznik","naglowki_okresow","sekcje","kapital_obrotowy_z_bilansu","uwagi"],
      SPOLKI, [10,14,46,16,12,14,10,40,60,24,60])

ws = wb.create_sheet("README")
for line in [
 ["SLOWNIK MAPOWANIA SPRAWOZDAN -> STANDARD BIZNESRADAR"],[""],
 ["Jak to dziala"],
 ["1. Arkusz 'Klucze' to slownik docelowy - klucze standardowe wyciagniete z NEU.xlsx (arkusz Spr Fin, kol. A i B)."],
 ["2. Arkusz 'Aliasy' to serce slownika: kazda linia to regula 'ta nazwa w tej sekcji -> ten klucz standardowy'."],
 ["   - pusta kolumna 'spolka' = alias GLOBALNY, dziala dla wszystkich spolek (tak rosnie slownik)"],
 ["   - wpisana spolka = wyjatek tylko dla niej, nadpisuje alias globalny"],
 ["   - 'sekcja' rozroznia te same nazwy w roznych miejscach (np. 'Udzielone pozyczki' w aktywach trwalych i obrotowych)"],
 ["   - 'sekcja' = * oznacza dowolna sekcje"],
 ["   - AGREGACJA: kilka wierszy z tym samym standard_key sumuje sie (np. 4 linie -> noncurrent_investments)"],
 ["   - ROZDZIELENIE: ten sam wiersz zrodlowy moze miec 2 aliasy (np. srodki pieniezne -> cash ORAZ current_investments)"],
 ["   - 'znak' = -1 gdy spolka podaje koszt na minusie, a standard oczekuje wartosci dodatniej"],
 ["   - standard_key = POMIN oznacza swiadome pominiecie (sumy posrednie, naglowki, pozycje poza standardem)"],
 ["3. Arkusz 'Reguly_obliczane' - klucze, ktorych nie ma wprost w raporcie, liczone z innych kluczy."],
 ["4. Arkusz 'Spolki' - konfiguracja parsera per plik (arkusz, kolumny, naglowki sekcji)."],[""],
 ["Uruchomienie"],
 ["python3 map.py --slownik slownik.xlsx --spolka NEU --plik \"Neuca za I kwartal 2026 roku_sprawozdania.xlsx\" --out wynik_NEU.xlsx"],[""],
 ["Wynik zawiera arkusze: Wynik (kolumna w standardzie), Walidacja (14 sum kontrolnych), Audyt (slad kazdego wiersza), Niezmapowane."],[""],
 ["Dodanie kolejnej spolki"],
 ["1. Dopisz wiersz w 'Spolki' (arkusz, kolumny, naglowki sekcji z jej pliku)."],
 ["2. Odpal map.py - pozycje o nazwach juz znanych zmapuja sie same z globalnych aliasow."],
 ["3. Zajrzyj do arkusza 'Niezmapowane' + podpowiedzi fuzzy na konsoli, dopisz brakujace aliasy do 'Aliasy'."],
 ["4. Powtarzaj az 'Walidacja' bedzie caly czas OK. Z kazda spolka slownik lapie coraz wiecej za pierwszym razem."],[""],
 ["Kolory w arkuszu Aliasy: zolty = DO_WERYFIKACJI (decyzja merytoryczna), szary = POMIN."],[""],
 ["Metodologia potwierdzona wzorcami NEU.xlsx, KGH.xlsx i AOL.xlsx (I 2026)"],
 ["- change_in_other_assets = (inne zob. krotkoterminowe + rozliczenia miedzyokresowe) Q - (te same) Q-1"],
 ["- UKLAD USTAWY O RACHUNKOWOSCI (AOL): RZiS w wariancie POROWNAWCZYM - cale 'Koszty dzialalnosci"],
 ["  operacyjnej' ida w cost_of_sales, a distribution_expenses i administrative_expenses zostaja PUSTE."],
 ["- W UoR reckoning = 'Rezerwy na zobowiazania' + 'Rozliczenia miedzyokresowe' (pasywa)."],
 ["- W UoR year_profit JEST wypelniany (ustawa rozdziela wynik okresu od lat ubieglych)."],
 ["- Dane w zlotych: mnoznik 0.001, kazda pozycja zrodlowa zaokraglana do pelnych tysiecy."],
 ["  Przy takich spolkach uzywaj tolerancji 2 w walidacji i porownaniu (zaokraglenia daja +/-1)."],
 ["- Gdy kolumna porownawcza NIE jest poprzednim kwartalem (UoR podaje rok wczesniej), kapitalu"],
 ["  obrotowego nie da sie policzyc ze sprawozdania. Podaj bilans poprzedniego kwartalu z pliku wzorcowego:"],
 ["     prev = brmap.wczytaj_okres_wzorca('AOL.xlsx', '2025-12-31')"],
 ["     brmap.mapuj('slownik.xlsx', 'AOL', plik, bilans_poprzedni=prev)"],
 ["- KAPITAL OBROTOWY w przeplywach NIE jest przepisywany z rachunku przeplywow spolki."],
 ["  Liczy sie go z roznic kolejnych BILANSOW (flaga kapital_obrotowy_z_bilansu w arkuszu Spolki):"],
 ["     change_in_inventories  = zapasy(Q-1) - zapasy(Q)"],
 ["     change_in_receivables  = naleznosci_kr(Q-1) - naleznosci_kr(Q)"],
 ["     change_in_payables     = [zob_kr - zob_kr_inne](Q) - [zob_kr - zob_kr_inne](Q-1)"],
 ["     change_in_other_assets = zob_kr_inne(Q) - zob_kr_inne(Q-1)"],
 ["  Zgodnosc co do zlotowki na obu spolkach, 10/10 pozycji."],
 ["- current_borrowings / noncurrent_borrowings = WYLACZNIE linia 'Kredyty i pozyczki'."],
 ["  Faktoring i 'Inne zobowiazania finansowe' ida w *_other_liabilities."],
 ["- property zawiera prawo do uzytkowania; right_to_use_assets to pozycja 'w tym' (nie sumuje sie)."],
 ["- retained_earnings = CALA pozycja zyskow zatrzymanych; year_profit zostaje PUSTY w kwartalach."],
 ["- Kapitaly z wyceny/OCI nie wchodza do rozbicia kapitalu wlasnego -> klucz pomocniczy"],
 ["  kapitaly_pozostale_nierozdzielone (nie trafia do arkusza Wynik, ale domyka sume kontrolna)."],[""],
 ["Uwagi techniczne"],
 ["- Nazwy sa porownywane po normalizacji: male litery, bez polskich znakow i interpunkcji."],
 ["  Dlatego 'ZYSK NETTO' i 'Zysk netto' to ten sam alias - nie trzeba wpisywac obu wariantow zapisu."],
 ["  Gdyby jednak trafily oba, skrypt je odfiltruje i wypisze ostrzezenie (bez tego pozycja liczylaby sie podwojnie)."],
 ["- Kolumna 'mnoznik' w arkuszu Spolki sprowadza jednostki do tysiecy (KGHM raportuje w mln -> 1000)."],
 ["- Znaczniki sekcji: 'Naglowek=>KOD' otwiera blok, '$Suma=>KOD' zamyka blok (dla spolek bez naglowkow)."],
 ["- Reguly z 'tylko_gdy_brak' = TAK licza sie wylacznie gdy spolka nie poda pozycji wprost."],
 ["- Ostatnia kontrola (stopa podatku) jest MIEKKA: 'PODEJRZANA' to sygnal do sprawdzenia, nie blad."],
]:
    ws.append(line)
ws.column_dimensions["A"].width = 140
ws["A1"].font = Font(bold=True, size=14)
for r in (3, 16, 20):
    ws.cell(row=r, column=1).font = Font(bold=True)
wb.save(OUT)
print("zapisano", OUT, "| kluczy:", len(klucze), "| aliasow:", len(A))

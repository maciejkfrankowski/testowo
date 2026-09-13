# -*- coding: utf-8 -*-
"""
brmap - mapowanie sprawozdan spolek na standard biznesradar.

Dziala tak samo lokalnie i w Google Colab:

    import brmap
    brmap.nowa_spolka("raport.xlsx", "slownik.xlsx", "ABC")   # kreator: konfiguracja + propozycje aliasow
    brmap.mapuj("slownik.xlsx", "ABC", "raport.xlsx")         # wlasciwe mapowanie
    brmap.porownaj("ABC.xlsx", "2026-03-31", "wynik_ABC.xlsx")# kontrola ze wzorcem
"""
import argparse, re, unicodedata, datetime, os
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ---------- normalizacja ----------
def norm(s):
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = s.replace("–", "-").replace("—", "-").replace("’", "'")
    myslnik = s.startswith("-")
    # UWAGA: 'l' z kreska (U+0142) NIE jest znakiem skladanym, wiec NFKD go nie rozlozy.
    # Bez tego "Przeplywy" normalizowalo sie do "przep ywy" i psulo dopasowania.
    # "Ŝ" to pozostalosc po starym kodowaniu (Mazovia/CP1250) - w polskich sprawozdaniach
    # zawsze oznacza "z": "NaleŜnosci", "PoŜyczki", "pienięŜne" (tak ma Dom Development).
    for a, b in (("ł", "l"), ("Ł", "L"), ("Ŝ", "z"), ("ŝ", "z"),
                 ("đ", "d"), ("ø", "o"), ("æ", "ae"), ("ß", "ss")):
        s = s.replace(a, b)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    # wiodacy myslnik oznacza pozycje 'w tym' i MUSI przetrwac normalizacje,
    # inaczej "- udzielone pozyczki" zlewa sie z "Udzielone pozyczki" i pozycja liczy sie dwa razy
    return ("- " + s) if myslnik else s

def similar(a, b):
    """prosty wspolczynnik podobienstwa 0-1 na bigramach slow (bez zaleznosci zewn.)"""
    A, B = set(a.split()), set(b.split())
    if not A or not B:
        return 0.0
    return 2 * len(A & B) / (len(A) + len(B))

# ---------- wczytanie slownika ----------
def read_sheet(wb, name):
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(h).strip() if h else "" for h in rows[0]]
    out = []
    for r in rows[1:]:
        if all(c is None or str(c).strip() == "" for c in r):
            continue
        out.append({hdr[i]: r[i] for i in range(len(hdr))})
    return out

def load_slownik(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    return {
        "klucze":   read_sheet(wb, "Klucze"),
        # drugi szablon standardu - banki maja wlasny zestaw kluczy (RAP_B)
        "klucze_b": read_sheet(wb, "Klucze_B") if "Klucze_B" in wb.sheetnames else [],
        "aliasy":   read_sheet(wb, "Aliasy"),
        "reguly":   read_sheet(wb, "Reguly_obliczane"),
        "spolki":   read_sheet(wb, "Spolki"),
        # arkusz opcjonalny: niezweryfikowane nazwy uzywane WYLACZNIE do podpowiedzi kreatora.
        # apply_mapping ich nie widzi - do mapowania licza sie tylko 'Aliasy'.
        "kandydaci": read_sheet(wb, "Aliasy_kandydaci") if "Aliasy_kandydaci" in wb.sheetnames else [],
    }

# ---------- parsowanie pliku spolki ----------
def parse_company(path, cfg):
    """Zwraca liste dictow: {sekcja, wiersz, nazwa, nazwa_norm, wartosci:[...]}

    Znaczniki sekcji (kolumna 'sekcje' w arkuszu Spolki):
      "Naglowek=>KOD"   - znacznik POCZATKU: wiersze PO nim naleza do KOD
      "$Suma=>KOD"      - znacznik KONCA: wiersze PRZED nim (i on sam) naleza do KOD
    Drugi wariant jest potrzebny dla spolek, ktore nie daja naglowkow blokow,
    tylko sumy na koncu (np. KGHM: '... Aktywa trwale 46696' zamyka blok).
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[cfg["arkusz"]]
    # kol_nazwa moze byc lista, np. "1,2". Sprawozdania z ESEF miewaja tabele
    # dwupoziomowa: kolumna 1 to strona bilansu ("Aktywa"/"Pasywa"), a nazwa pozycji
    # stoi dopiero w kolumnie 2. Bierzemy wtedy OSTATNIA niepusta komorke, ktora nie
    # jest liczba - w rachunku wynikow kolumna 2 trzyma juz wartosc, wiec odpada sama.
    cols_name = [int(c) for c in str(cfg["kol_nazwa"]).split(",") if str(c).strip()] or [1]
    cols_val = [int(c) for c in str(cfg["kol_wartosci"]).split(",")]
    # Niektore spolki podaja rachunek wynikow narastajaco w innej kolumnie niz bilans
    # (Wawel za 9 miesiecy: RZiS w kol. 3, bilans na 30.09 w kol. 2). Wtedy 'kol_wartosci_bilans'
    # nadpisuje kolumny dla sekcji bilansowych.
    cols_bil = [int(c) for c in str(cfg.get("kol_wartosci_bilans") or "").split(",") if str(c).strip()]
    # ... a rachunek wynikow jeszcze inne (Pekao: RZiS ma 4 kolumny, przeplywy 2)
    cols_rzis = [int(c) for c in str(cfg.get("kol_wartosci_rzis") or "").split(",") if str(c).strip()]
    mnoznik = float(cfg.get("mnoznik") or 1)

    start_mark, end_mark = {}, {}
    for pair in str(cfg["sekcje"]).split("|"):
        if not pair.strip():
            continue
        k, v = pair.split("=>")
        k = k.strip()
        (end_mark if k.startswith("$") else start_mark)[_bez_numeracji(k.lstrip("$"))] = v.strip()

    SEKCJE_BILANSU = {"BILANS", "AKT_TRW", "AKT_OBR", "AKTYWA", "KAPITAL", "ZOB_DL", "ZOB_KR", "PASYWA",
                      "BANK_A", "BANK_P"}
    SEKCJE_RZIS = {"RZIS", "RZIS_ZYSK", "RZIS_CD"}
    items, bufor, sekcja_biezaca = [], [], "?"

    def flush(sekcja):
        for it in bufor:
            it["sekcja"] = sekcja
            if cols_bil and sekcja in SEKCJE_BILANSU:
                it["wartosci"] = it["wartosci_bilans"]
            elif cols_rzis and sekcja in SEKCJE_RZIS:
                it["wartosci"] = it["wartosci_rzis"]
            it.pop("wartosci_bilans", None)
            it.pop("wartosci_rzis", None)
            items.append(it)
        bufor.clear()

    for i, row in enumerate(ws.iter_rows(min_row=1, values_only=True), 1):
        nazwa = None
        for _c in cols_name:
            v = row[_c - 1] if len(row) >= _c else None
            if v is None or str(v).strip() == "" or isinstance(v, (int, float)):
                continue
            nazwa = str(v).strip()
        if nazwa is None:
            continue
        n = _bez_numeracji(nazwa)

        def czytaj(kolumny):
            out = []
            for c in kolumny:
                v = row[c - 1] if len(row) >= c else None
                if isinstance(v, (int, float)):
                    v = v * mnoznik
                    out.append(round(v) if mnoznik < 1 else v)   # dane w zlotych -> pelne tysiace
                else:
                    out.append(None)
            return out

        vals = czytaj(cols_val)
        it = {"sekcja": None, "wiersz": i, "nazwa": str(nazwa).strip(),
              "nazwa_norm": norm(nazwa), "wartosci": vals,
              "wartosci_bilans": czytaj(cols_bil) if cols_bil else vals,
              "wartosci_rzis": czytaj(cols_rzis) if cols_rzis else vals}

        if n in start_mark:                       # naglowek otwierajacy blok
            flush(sekcja_biezaca)
            sekcja_biezaca = start_mark[n]
            if not all(v is None for v in vals):  # naglowek z liczbami - zalicz do nowej sekcji
                bufor.append(it)
            continue
        if n in end_mark:                         # suma zamykajaca blok
            kod = end_mark[n]
            # "Aktywa razem" / "Pasywa razem" to pojedyncze linie zbiorcze: nie zabieraja
            # bufora, jesli wewnatrz trwa juz drobniejsza sekcja (uklad UoR)
            if kod in ("AKTYWA", "PASYWA") and sekcja_biezaca not in ("AKTYWA", "PASYWA", "?") and bufor:
                flush(sekcja_biezaca)
                bufor.append(it)        # przez flush, zeby zadzialala podmiana kolumn bilansu
                flush(kod)
            else:
                bufor.append(it)
                flush(kod)
            continue
        bufor.append(it)
    flush(sekcja_biezaca)
    return items

DATA_RE = re.compile(r"(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})")


def daty_z_naglowkow(naglowki):
    """Wyciaga daty bilansowe z opisow okresow (ostatnia data w kazdym opisie)."""
    out = []
    for h in str(naglowki or "").split("|"):
        znal = DATA_RE.findall(h)
        if not znal:
            out.append(None)
            continue
        d, m, y = znal[-1]
        y = int(y) + (2000 if int(y) < 100 else 0)
        try:
            out.append(datetime.datetime(y, int(m), int(d)))
        except ValueError:
            out.append(None)
    return out


# ---------- mapowanie ----------
def build_index(aliasy, spolka):
    """(sekcja, nazwa_norm) -> lista regul. Reguly per-spolka biją globalne."""
    idx = defaultdict(list)
    for a in aliasy:
        sp = (a.get("spolka") or "").strip().upper()
        if sp and sp != spolka.upper():
            continue
        key = ((a.get("sekcja") or "").strip(), norm(a.get("nazwa_oryginalna")))
        idx[key].append(a)
    # per-spolka nadpisuje globalne
    out, kolizje = {}, []
    for k, v in idx.items():
        spec = [r for r in v if (r.get("spolka") or "").strip()]
        v = spec if spec else v
        # deduplikacja: identyczna regula wpisana dwa razy (np. dwa warianty zapisu tej samej
        # nazwy, ktore po normalizacji sa tozsame) liczylaby pozycje podwojnie
        widziane, unikalne = set(), []
        for r in v:
            sig = ((r.get("standard_key") or "").strip(), float(r.get("znak") or 1), float(r.get("waga") or 1))
            if sig in widziane:
                kolizje.append((k[0], k[1], sig[0]))
                continue
            widziane.add(sig)
            unikalne.append(r)
        out[k] = unikalne
    if kolizje:
        print("UWAGA - zdublowane reguly w slowniku (pominieto powtorzenia, inaczej byloby podwojne liczenie):")
        for sek, nazwa, sk in kolizje:
            print(f"   [{sek}] {nazwa} -> {sk}")
    return out

def apply_mapping(items, idx, n_kol):
    wyniki = defaultdict(lambda: [0.0] * n_kol)
    uzyte = defaultdict(list)
    audyt, niezmapowane = [], []
    for it in items:
        key = (it["sekcja"], it["nazwa_norm"])
        reguly = idx.get(key)
        if not reguly:
            reguly = idx.get(("*", it["nazwa_norm"]))       # alias globalny bez sekcji
        if not reguly:
            niezmapowane.append(it)
            continue
        for r in reguly:
            sk = (r.get("standard_key") or "").strip()
            if not sk or sk.upper() in ("POMIN", "IGNORUJ"):
                audyt.append({**it, "standard_key": "(pominieto)", "znak": "",
                              "pewnosc": r.get("pewnosc"), "uwagi": r.get("uwagi")})
                continue
            znak = float(r.get("znak") or 1)
            waga = float(r.get("waga") or 1)
            for j, v in enumerate(it["wartosci"]):
                if v is not None:
                    wyniki[sk][j] += znak * waga * v
            uzyte[sk].append(it["nazwa"])
            audyt.append({**it, "standard_key": sk, "znak": znak,
                          "pewnosc": r.get("pewnosc"), "uwagi": r.get("uwagi")})
    return wyniki, uzyte, audyt, niezmapowane

def apply_reguly(wyniki, reguly, n_kol):
    """Klucze wyliczane. Formula w notacji: klucz +/- klucz ..."""
    for r in sorted(reguly, key=lambda x: int(x.get("kolejnosc") or 99)):
        sk = (r.get("standard_key") or "").strip()
        formula = str(r.get("formula") or "").strip()
        if not sk or not formula:
            continue
        # regula warunkowa: licz tylko jesli spolka nie podala tej pozycji wprost
        if str(r.get("tylko_gdy_brak") or "").strip().upper() in ("TAK", "T", "1", "YES"):
            if sk in wyniki and any(abs(v) > 0 for v in wyniki[sk]):
                continue
        toks = re.findall(r"([+-]?)\s*([a-z_0-9]+)", formula)
        vals = [0.0] * n_kol
        for sign, name in toks:
            s = -1.0 if sign == "-" else 1.0
            src = wyniki.get(name, [0.0] * n_kol)
            for j in range(n_kol):
                vals[j] += s * (src[j] or 0.0)
        wyniki[sk] = vals
    return wyniki

def kapital_obrotowy_z_bilansu(w, n_kol, poprzedni=None):
    """Odtwarza metodologie biznesradar: zmiany kapitalu obrotowego NIE sa przepisywane
    z rachunku przeplywow spolki, tylko liczone jako roznice kolejnych BILANSOW.
    Potwierdzone co do zlotowki na NEU.xlsx, KGH.xlsx i AOL.xlsx (I 2026).

    'poprzedni' = slownik {standard_key: wartosc} z bilansu poprzedniego kwartalu,
    potrzebny gdy sprawozdanie spolki nie zawiera go w drugiej kolumnie (uklad UoR
    podaje ten sam kwartal rok wczesniej).
    """
    KLUCZE = ("change_in_inventories", "change_in_receivables", "change_in_payables",
              "change_in_other_assets", "change_in_working_capital")
    for k in KLUCZE:
        w.setdefault(k, [0.0] * n_kol)

    def biez(k, j):
        v = w.get(k)
        x = (v[j] if v and j < len(v) else 0) or 0
        return x if isinstance(x, (int, float)) else 0

    for j in range(n_kol):
        if j == 0 and poprzedni is not None:
            odn = lambda k: poprzedni.get(k) or 0          # okres odniesienia z pliku wzorcowego
        elif j + 1 < n_kol:
            odn = lambda k, _j=j: biez(k, _j + 1)          # druga kolumna sprawozdania
        else:
            for k in KLUCZE:
                w[k][j] = None
            continue
        teraz = lambda k, _j=j: biez(k, _j)
        dlug  = lambda f: f("current_liabilities") - f("current_other_liabilities")
        inne  = lambda f: f("current_other_liabilities") + f("reckoning")
        w["change_in_inventories"][j] = odn("inventory") - teraz("inventory")
        w["change_in_receivables"][j] = odn("current_receivables") - teraz("current_receivables")
        w["change_in_payables"][j] = dlug(teraz) - dlug(odn)
        w["change_in_other_assets"][j] = inne(teraz) - inne(odn)
        w["change_in_working_capital"][j] = sum(w[k][j] for k in KLUCZE[:4])
    return w

# ---------- walidacja ----------
def g_pub(k, j, w, n_kol):
    x = (w.get(k) or [0.0] * n_kol)[j] or 0.0
    return x if isinstance(x, (int, float)) else 0.0

def waliduj_bank(w, n_kol, tol=1.0):
    def g(k, j):
        x = (w.get(k) or [0.0] * n_kol)[j] or 0.0
        return x if isinstance(x, (int, float)) else 0.0
    AKT = ["cash_and_balances_with_the_central_bank", "loans_to_banks", "financial_assets_held_for_trading",
           "held_to_maturity_financial_assets", "loans_to_customers", "other_financial_assets",
           "investments_in_subsidiaries", "intangible_assets", "property", "tax_assets", "other_assets"]
    ZOB = ["liabilities_due_to_the_central_bank", "liabilities_due_to_banks",
           "held_to_maturity_financial_liabilities", "financial_liabilities_held_for_trading",
           "liabilities_due_to_customers", "debt_securities", "subordinated_liabilities",
           "tax_liabilities", "other_liabilities"]
    KAP = ["share_capital", "share_premium", "own_share", "revaluation_reserves",
           "retained_earnings", "year_profit"]
    OPER = ["net_interest_income", "net_fee_income", "dividend_income", "realised_gains",
            "net_other_financial_income", "net_other_operating_income", "impairment_losses",
            "administration_costs", "net_other_operating_costs", "bank_tax"]
    testy = [
        ("Aktywa = Pasywa", lambda j: g("total_assets", j) - g("total_equity_liabilities", j)),
        ("Skladniki aktywow", lambda j: sum(g(k, j) for k in AKT) - g("total_assets", j)),
        ("Skladniki zobowiazan", lambda j: sum(g(k, j) for k in ZOB) - g("total_liabilities", j)),
        ("Kapital wlasny akcjonariuszy", lambda j: sum(g(k, j) for k in KAP) - g("own_capital", j)),
        ("Kapitaly razem", lambda j: g("own_capital", j) + g("non_control_profit", j) - g("total_capital", j)),
        ("Zobowiazania + kapitaly = pasywa",
         lambda j: g("total_liabilities", j) + g("total_capital", j) - g("total_equity_liabilities", j)),
        ("WNiP = firma + pozostale",
         lambda j: g("goodwill", j) + g("other_intangible_assets", j) - g("intangible_assets", j)),
        ("Wynik odsetkowy", lambda j: g("interest_income", j) - g("interest_expense", j) - g("net_interest_income", j)),
        ("Wynik prowizyjny", lambda j: g("fee_income", j) - g("fee_expense", j) - g("net_fee_income", j)),
        ("Wynik operacyjny", lambda j: sum(g(k, j) for k in OPER) - g("net_operating_profit", j)),
        ("Zysk przed opodatkowaniem",
         lambda j: g("net_operating_profit", j) + g("related_income", j) - g("beforetax_profit", j)),
        ("Przeplywy razem", lambda j: g("operating_cashflow", j) + g("investing_cashflow", j)
         + g("financing_cashflow", j) - g("net_cashflow", j)),
    ]
    out = []
    for nazwa, fn in testy:
        row = [nazwa]
        for j in range(n_kol):
            d = fn(j)
            row.append("OK" if abs(d) <= tol else f"ROZNICA {d:,.0f}")
        out.append(row)
    row = ["Implikowana stopa podatku (0-40%)"]
    for j in range(n_kol):
        bt, npf = g("beforetax_profit", j), g("net_profit", j)
        r = (bt - npf) / bt if bt > 0 else None
        row.append("brak danych" if r is None else
                   (f"OK ({r:.1%})" if 0 <= r <= 0.40 else f"PODEJRZANA {r:.1%}"))
    out.append(row)
    return out


def waliduj(w, n_kol, tol=1.0):
    def g(k, j):
        x = (w.get(k) or [0.0] * n_kol)[j] or 0.0
        return x if isinstance(x, (int, float)) else 0.0
    testy = [
        ("Aktywa = Pasywa", lambda j: g("total_assets", j) - g("total_equity_liabilities", j)),
        # assets_for_sale to linia INFORMACYJNA wewnatrz aktywow obrotowych, nie trzeci skladnik sumy
        ("Aktywa trwale + obrotowe = suma", lambda j: g("noncurrent_assets", j) + g("current_assets", j) - g("total_assets", j)),
        ("Skladniki aktywow trwalych", lambda j: g("property", j) + g("intangible_assets", j) + g("noncurrent_investments", j) + g("noncurrent_receivables", j) + g("other_noncurrent_assets", j) - g("noncurrent_assets", j)),
        ("Skladniki aktywow obrotowych", lambda j: g("inventory", j) + g("current_receivables", j) + g("current_investments", j) + g("other_current_assets", j) + g("assets_for_sale", j) - g("current_assets", j)),
        ("Skladniki kapitalu wlasnego", lambda j: g("share_capital", j) + g("own_share", j) + g("reserve", j) + g("retained_earnings", j) + g("year_profit", j) + g("nonshare_capital", j) + g("kapitaly_pozostale_nierozdzielone", j) - g("capital", j)),
        ("Kapital + zobowiazania = pasywa", lambda j: g("capital", j) + g("noncurrent_liabilities", j) + g("current_liabilities", j) + g("reckoning", j) - g("total_equity_liabilities", j)),
        ("Skladniki zob. dlugoterminowych", lambda j: g("noncurrent_trade_payables", j) + g("noncurrent_borrowings", j) + g("noncurrent_obligations", j) + g("noncurrent_leasing", j) + g("noncurrent_other_liabilities", j) - g("noncurrent_liabilities", j)),
        ("Skladniki zob. krotkoterminowych", lambda j: g("current_trade_payables", j) + g("current_borrowings", j) + g("current_obligations", j) + g("current_leasing", j) + g("current_other_liabilities", j) - g("current_liabilities", j)),
        ("Zysk ze sprzedazy", lambda j: g("revenues", j) - g("cost_of_sales", j) - g("distribution_expenses", j) - g("administrative_expenses", j) - g("gross_profit", j)),
        ("EBIT", lambda j: g("gross_profit", j) + g("other_operating_income", j) - g("other_operating_costs", j) - g("ebit", j)),
        ("Zysk przed opodatkowaniem", lambda j: g("ebit", j) + g("finance_income", j) - g("finance_costs", j) + g("other_income", j) - g("before_tax_profit", j)),
        ("Przeplywy razem", lambda j: g("operating_cashflow", j) + g("investing_cashflow", j) + g("financing_cashflow", j) - g("net_cashflow", j)),
        ("Kapital obrotowy = suma skladnikow", lambda j: g("change_in_receivables", j) + g("change_in_inventories", j) + g("change_in_payables", j) + g("change_in_other_assets", j) - g("change_in_working_capital", j)),
    ]
    # kontrola miekka: implikowana stopa podatku poza 0-40% zwykle oznacza blad mapowania wyniku
    def stopa(j):
        bt, npf = g_pub("before_tax_profit", j, w, n_kol), g_pub("net_profit", j, w, n_kol)
        if bt <= 0:
            return None
        return (bt - npf) / bt

    out = []
    for nazwa, fn in testy:
        row = [nazwa]
        for j in range(n_kol):
            d = fn(j)
            row.append("OK" if abs(d) <= tol else f"ROZNICA {d:,.0f}")
        out.append(row)
    row = ["Implikowana stopa podatku (0-40%)"]
    for j in range(n_kol):
        r = stopa(j)
        row.append("brak danych" if r is None else
                   (f"OK ({r:.1%})" if 0 <= r <= 0.40 else f"PODEJRZANA {r:.1%}"))
    out.append(row)
    return out

# ---------- zapis ----------
BOLD = Font(bold=True)
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FILL = PatternFill("solid", fgColor="C6EFCE")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")

def style_header(ws, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"

def zapisz(out_path, klucze, wyniki, naglowki, audyt, niezmapowane, walid):
    wb = openpyxl.Workbook()
    n_kol = len(naglowki)

    ws = wb.active; ws.title = "Wynik"
    ws.append(["standard_key", "etykieta_pl", "sekcja"] + naglowki)
    for k in klucze:
        sk = (k.get("standard_key") or "").strip()
        if not sk:
            continue
        v = wyniki.get(sk)
        if v is None:   # klucz bez zrodla: 0 dla pozycji finansowych, pusto dla metadanych
            v = [None] * n_kol if (k.get("sekcja") or "") == "META" else [0.0] * n_kol
        ws.append([sk, k.get("etykieta_pl"), k.get("sekcja")] +
                  [(round(x, 2) if isinstance(x, (int, float)) else x) for x in v])
    style_header(ws, 3 + n_kol)
    ws.column_dimensions["A"].width = 30; ws.column_dimensions["B"].width = 48
    ws.column_dimensions["C"].width = 16
    for c in range(4, 4 + n_kol):
        ws.column_dimensions[get_column_letter(c)].width = 18
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=c).number_format = "#,##0"

    ws = wb.create_sheet("Walidacja")
    ws.append(["test"] + naglowki)
    for row in walid:
        ws.append(row)
        for c in range(2, 2 + n_kol):
            cell = ws.cell(row=ws.max_row, column=c)
            v = str(cell.value)
            cell.fill = OK_FILL if v.startswith("OK") else (WARN_FILL if v == "brak danych" else BAD_FILL)
    style_header(ws, 1 + n_kol); ws.column_dimensions["A"].width = 40
    for c in range(2, 2 + n_kol):
        ws.column_dimensions[get_column_letter(c)].width = 20

    ws = wb.create_sheet("Audyt")
    ws.append(["wiersz", "sekcja", "nazwa w raporcie", "standard_key", "znak", "pewnosc", "uwagi"] + naglowki)
    for a in audyt:
        ws.append([a["wiersz"], a["sekcja"], a["nazwa"], a["standard_key"], a["znak"],
                   a.get("pewnosc"), a.get("uwagi")] + a["wartosci"])
        if str(a.get("pewnosc") or "").upper().startswith("DO_WER"):
            for c in range(1, 8):
                ws.cell(row=ws.max_row, column=c).fill = WARN_FILL
    style_header(ws, 7 + n_kol)
    for col, wdt in zip("ABCDEFG", (8, 22, 62, 30, 7, 16, 46)):
        ws.column_dimensions[col].width = wdt

    ws = wb.create_sheet("Niezmapowane")
    ws.append(["wiersz", "sekcja", "nazwa w raporcie"] + naglowki)
    for it in niezmapowane:
        ws.append([it["wiersz"], it["sekcja"], it["nazwa"]] + it["wartosci"])
    style_header(ws, 3 + n_kol)
    for col, wdt in zip("ABC", (8, 22, 78)):
        ws.column_dimensions[col].width = wdt
    wb.save(out_path)


# =====================================================================
#  KREATOR NOWEJ SPOLKI - automatyczne wykrycie ukladu pliku
# =====================================================================

# wzorce naglowkow blokow. Kolejnosc ma znaczenie - pierwszy trafiony wygrywa.
WZORCE_SEKCJI = [
    (r"^(income statement|rachunek zyskow|statement of comprehensive income)", "RZIS"),
    (r"^(cash flow statement|rachunek przeplywow)", "CF"),
    (r"^(balance sheet|statement of financial position)", "BILANS"),
    (r"^bilans$", "BILANS"),
    (r"^przeplywy (srodkow )?(pienieznych|pieniezne)( netto)? z dzialal.*operacyjn", "CF"),
    (r"^przeplywy (srodkow )?(pienieznych|pieniezne)( netto)? z dzialal.*inwest", "CF_INW"),
    (r"^przeplywy (srodkow )?(pienieznych|pieniezne)( netto)? z dzialal.*finansow", "CF_FIN"),
    (r"^aktywa trwale( razem| ogolem)?$", "AKT_TRW"),
    (r"^aktywa obrotowe( razem| ogolem)?$", "AKT_OBR"),
    (r"^(kapital wlasny|kapitaly wlasne)( razem| ogolem)?$", "KAPITAL"),
    (r"^zobowiazania dlugoterminowe( razem| ogolem)?$", "ZOB_DL"),
    (r"^zobowiazania krotkoterminowe( razem| ogolem)?$", "ZOB_KR"),
    (r"^(razem |suma )?aktywa(ow)?( razem| ogolem)?$", "AKTYWA"),
    (r"^suma aktywow$", "AKTYWA"),
    (r"^(razem |suma )?(pasywa|pasywow|zobowiazania i kapital wlasny|kapital wlasny i zobowiazania)( razem| ogolem)?$", "PASYWA"),
    (r"^zobowiazania i rezerwy na zobowiazania$", "PASYWA"),
]

# numeracja pozycji w ustawie o rachunkowosci: "A.", "I.", "1.", "a)", "-"
RE_NUMERACJA = re.compile(r"^([A-Z]|[IVXLC]+|\d+)[\.\)]\s+|^[a-z]\)\s+|^-\s+")

def _bez_numeracji(nazwa):
    return norm(RE_NUMERACJA.sub("", str(nazwa or "").strip()))

def _ma_numeracje(nazwa):
    """Pozycje numerowane (uklad UoR) podaja sume NA POCZATKU bloku, nie na koncu."""
    return bool(re.match(r"^([A-Z]|[IVXLC]+|\d+)[\.\)]\s+", str(nazwa or "").strip()))

def _dopasuj_sekcje(n):
    for wzor, kod in WZORCE_SEKCJI:
        if re.match(wzor, n):
            return kod
    return None

def zbadaj_plik(plik, arkusz=None):
    """Czyta nieznany plik sprawozdania i proponuje konfiguracje do arkusza 'Spolki'.

    Kluczowa heurystyka: jesli wiersz-naglowek sekcji MA liczby, to jest suma
    ZAMYKAJACA blok (znacznik $). Jesli nie ma liczb - OTWIERA blok.
    Dokladnie tym rozni sie uklad KGHM od ukladu Neuki.
    """
    wb = openpyxl.load_workbook(plik, data_only=True)
    arkusz = arkusz or wb.sheetnames[0]
    ws = wb[arkusz]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    szer = max((len(r) for r in rows), default=0)

    # kolumna z nazwami = ta z najwieksza liczba tekstow
    ile_tekst = [sum(1 for r in rows if len(r) > c and isinstance(r[c], str) and r[c].strip()) for c in range(szer)]
    kol_nazwa = (ile_tekst.index(max(ile_tekst)) + 1) if szer else 1

    # kolumny wartosci = te z liczbami (poza kolumna nazw)
    ile_liczb = [sum(1 for r in rows if len(r) > c and isinstance(r[c], (int, float))) for c in range(szer)]
    prog = max(ile_liczb) * 0.3 if ile_liczb and max(ile_liczb) else 0
    kol_wartosci = [c + 1 for c in range(szer) if c + 1 != kol_nazwa and ile_liczb[c] >= prog and ile_liczb[c] > 3]

    # jednostka. Grosze w wartosciach = dane w ZLOTYCH (standard trzyma tysiace).
    liczby = [v for r in rows for v in r if isinstance(v, (int, float)) and v]
    maks = max((abs(v) for v in liczby), default=0)
    z_groszami = sum(1 for v in liczby if abs(v - round(v)) > 1e-9)
    if liczby and z_groszami / len(liczby) > 0.3:
        mnoznik = 0.001
    elif maks < 1_000_000:
        mnoznik = 1000
    else:
        mnoznik = 1

    # naglowki okresow - ostatni wiersz 'Pozycja' (zwykle nad bilansem, ma daty)
    naglowki = []
    for r in rows:
        if len(r) > kol_nazwa - 1 and norm(r[kol_nazwa - 1]) == "pozycja":
            kand = [str(r[c - 1]) for c in kol_wartosci if len(r) >= c and r[c - 1] is not None]
            if len(kand) == len(kol_wartosci):
                naglowki = kand

    # --- znaczniki sekcji ---
    # Najpierw znajdz kandydatow, potem rozstrzygnij czy suma OTWIERA czy ZAMYKA blok.
    kandydaci = []
    for i, r in enumerate(rows, 1):
        nazwa = r[kol_nazwa - 1] if len(r) >= kol_nazwa else None
        if not isinstance(nazwa, str) or not nazwa.strip():
            continue
        kod = _dopasuj_sekcje(_bez_numeracji(nazwa))
        if kod:
            kandydaci.append((i, nazwa.strip(), kod))

    kol1 = kol_wartosci[0] if kol_wartosci else 2

    def wartosc(i):
        r = rows[i - 1]
        v = r[kol1 - 1] if len(r) >= kol1 else None
        return v if isinstance(v, (int, float)) else None

    def skladnik(i):
        """Czy wiersz jest samodzielna pozycja, a nie rozwinieciem 'w tym'."""
        r = rows[i - 1]
        n = r[kol_nazwa - 1] if len(r) >= kol_nazwa else None
        if not isinstance(n, str) or not n.strip():
            return False
        t = n.strip()
        return not (t[:1].islower() or t.startswith("-") or "w tym" in t.lower())

    def suma_do_przodu(poz, i, v):
        """Czy ktorykolwiek POCZATKOWY fragment wierszy PO sumie daje dokladnie te sume?
        Jesli tak, suma otwiera blok. Sprawdzamy narastajaco, bo dalej moga stac
        zagniezdzone sumy posrednie, ktore by wynik zdublowaly."""
        nastepny = kandydaci[poz + 1][0] if poz + 1 < len(kandydaci) else len(rows) + 1
        biezaca = 0.0
        for j in range(i + 1, nastepny):
            if not skladnik(j):
                continue
            biezaca += wartosc(j) or 0
            if abs(biezaca - v) <= max(2.0, abs(v) * 1e-5):
                return True
        return False

    # Uklad sprawozdania jest jednolity w calym pliku, wiec zamiast decydowac osobno
    # dla kazdej sumy - glosujemy. Pojedynczy blok potrafi zmylic (zagniezdzone sumy),
    # ale caly plik juz nie.
    glosy_poczatek = glosy_koniec = 0
    for poz, (i, nazwa, kod) in enumerate(kandydaci):
        v = wartosc(i)
        if v is None or _ma_numeracje(nazwa) or kod in ("AKTYWA", "PASYWA", "BILANS", "RZIS", "CF", "CF_INW", "CF_FIN"):
            continue
        if suma_do_przodu(poz, i, v):
            glosy_poczatek += 1
        else:
            glosy_koniec += 1
    sumy_otwieraja = glosy_poczatek > glosy_koniec

    sekcje, wykryte = [], []
    for poz, (i, nazwa, kod) in enumerate(kandydaci):
        v = wartosc(i)
        if v is None or _ma_numeracje(nazwa):
            typ = "poczatek"
        elif kod in ("AKTYWA", "PASYWA"):
            typ = "koniec"          # "Aktywa razem" / "Pasywa razem" to zawsze pojedyncza linia
        else:
            typ = "poczatek" if sumy_otwieraja else "koniec"
        sekcje.append(("$" if typ == "koniec" else "") + nazwa + "=>" + kod)
        wykryte.append({"wiersz": i, "nazwa": nazwa, "kod": kod, "typ": typ})

    return {
        "arkusz": arkusz,
        "kol_nazwa": kol_nazwa,
        "kol_wartosci": ",".join(map(str, kol_wartosci)),
        "mnoznik": mnoznik,
        "naglowki_okresow": "|".join(naglowki),
        "sekcje": "|".join(sekcje),
        "kapital_obrotowy_z_bilansu": "TAK",
        "wykryte_sekcje": wykryte,
        "maks_wartosc": maks,
    }

# klucze kosztowe - jesli spolka podaje je na minusie, standard chce wartosci dodatniej
KLUCZE_KOSZTOWE = {"cost_of_sales", "distribution_expenses", "administrative_expenses",
                   "other_operating_costs", "finance_costs", "capex",
                   "outflows_for_acquisitions", "repaid_bank_loans", "lease_liab_payments"}

# sumy zamykajace blok bilansu maja deterministyczny odpowiednik
TOTAL_SEKCJI = {"AKT_TRW": "noncurrent_assets", "AKT_OBR": "current_assets", "KAPITAL": "capital",
                "ZOB_DL": "noncurrent_liabilities", "ZOB_KR": "current_liabilities",
                "AKTYWA": "total_assets", "PASYWA": "total_equity_liabilities",
                "CF": "operating_cashflow", "CF_INW": "investing_cashflow", "CF_FIN": "financing_cashflow"}

def zaproponuj_aliasy(plik, cfg, slownik_path, spolka=""):
    """Dla pozycji spoza slownika proponuje standard_key. NIC nie zapisuje do slownika.

    Zasada nadrzedna: lepiej nie podpowiedziec nic niz podpowiedziec zle.
    Zla podpowiedz w zielonym kolorze zostanie klikniepta bez sprawdzenia - i to jest
    grozniejsze niz pusta komorka, ktora zmusza do decyzji.
    """
    sl = load_slownik(slownik_path)
    idx = build_index(sl["aliasy"], spolka or "___")
    sekcja_klucza = {(k.get("standard_key") or "").strip(): (k.get("sekcja") or "") for k in sl["klucze"]}
    klucze_std = [(k.get("standard_key"), k.get("etykieta_pl")) for k in sl["klucze"] if k.get("standard_key")]

    znane = []                       # (sekcja, nazwa_norm, standard_key, zweryfikowany)
    for (sek, nn), reguly in idx.items():
        for r in reguly:
            znane.append((sek, nn, (r.get("standard_key") or "").strip(), True))
    # Kandydaci (arkusz Aliasy_kandydaci) nie maja sekcji ani znaku i nie byli sprawdzeni
    # wzorcem. Pasuja do dowolnej sekcji, ale ich pewnosc jest z gory ograniczona do NISKA -
    # inaczej dopasowanie po dokladnej nazwie awansowaloby je na ZOLTE, ktore ma byc pewne.
    for k in sl.get("kandydaci", []):
        sk = str(k.get("standard_key") or "").strip()
        if sk and not sk.endswith("?"):
            znane.append(("*", norm(k.get("nazwa_oryginalna")), sk, False))

    # wiersze bedace znacznikami sekcji (sumy zamykajace bloki)
    markery = {s["wiersz"]: s for s in cfg.get("wykryte_sekcje", [])}

    items = parse_company(plik, cfg)
    propozycje = []
    for it in items:
        if idx.get((it["sekcja"], it["nazwa_norm"])) or idx.get(("*", it["nazwa_norm"])):
            continue                                        # juz w slowniku
        nazwa, wart = it["nazwa"], [v for v in it["wartosci"] if v is not None]
        n_tokenow = len(it["nazwa_norm"].split())
        prop, pewnosc, uwaga, kandydaci = "", "BRAK", "", []
        mark = markery.get(it["wiersz"])

        if mark and mark["typ"] == "koniec" and TOTAL_SEKCJI.get(mark["kod"]):
            prop, pewnosc = TOTAL_SEKCJI[mark["kod"]], "WYSOKA"
            uwaga = f"suma zamykajaca blok {mark['kod']} - odpowiednik jednoznaczny"
        elif nazwa[:1].islower():
            prop, pewnosc = "POMIN", "WYSOKA"
            uwaga = "pozycja 'w tym' (mala litera) - zawarta w wierszu nadrzednym, NIE liczyc drugi raz"
        elif not wart:
            prop, pewnosc, uwaga = "POMIN", "WYSOKA", "wiersz bez liczb - naglowek"
        else:
            # kandydaci: identyczna nazwa najpierw, potem rozmyte; ZAWSZE premia za te sama sekcje
            ranking = []
            for sek, nn, sk, pewny in znane:
                if not sk:
                    continue
                pod = 1.0 if nn == it["nazwa_norm"] else similar(it["nazwa_norm"], nn)
                if pod < 0.4:
                    continue
                premia = 0.35 if sek == it["sekcja"] else 0.0
                if not pewny:
                    premia -= 0.10          # zweryfikowany alias zawsze przed kandydatem
                ranking.append((pod + premia, pod, sek, nn, sk, pewny))
            ranking.sort(reverse=True)

            prog = 0.5 if n_tokenow >= 3 else 0.95     # krotkie nazwy typu "Pozostale" daja smieciowe dopasowania
            if ranking and ranking[0][1] >= prog:
                _, pod, sek, nn, sk, pewny = ranking[0]
                ta_sama = sek == it["sekcja"] and pewny
                # pozycja CF wskazujaca na klucz z RZiS/bilansu to zwykle powtorzenie punktu wyjscia
                if it["sekcja"].startswith("CF") and sekcja_klucza.get(sk, "") in ("RZIS", "BILANS_AKTYWA", "BILANS_PASYWA"):
                    prop = "POMIN"
                    pewnosc = "SREDNIA" if pewny else "NISKA"
                    uwaga = (f"w przeplywach powtarza sie pozycja z {sekcja_klucza.get(sk)} ({sk}) - zwykle duplikat"
                             + ("" if pewny else "; podpowiedz z arkusza kandydatow"))
                else:
                    prop = sk
                    if not pewny:
                        pewnosc = "NISKA"
                        uwaga = f"z arkusza kandydatow (niezweryfikowane, bez sekcji i znaku): {nn[:48]}"
                    elif pod >= 0.99 and ta_sama:
                        pewnosc, uwaga = "WYSOKA", "nazwa identyczna z istniejacym aliasem w tej samej sekcji"
                    elif pod >= 0.99:
                        pewnosc = "SREDNIA"
                        uwaga = f"nazwa identyczna, ale alias pochodzi z sekcji {sek} - SPRAWDZ czy tu znaczy to samo"
                    elif ta_sama and pod >= 0.85:
                        pewnosc, uwaga = "SREDNIA", f"dopasowanie {pod:.2f} w tej samej sekcji do: {nn[:52]}"
                    else:
                        pewnosc, uwaga = "NISKA", f"dopasowanie {pod:.2f} (sekcja {sek}) do: {nn[:52]}"
                kandydaci = [f"{p:.2f} {s}:{k}" for _, p, s, _, k, _w in ranking[1:4]]
            else:
                uwaga = "brak wiarygodnej podpowiedzi - zmapuj recznie (lista kluczy w arkuszu 4)"
                kandydaci = [f"{p:.2f} {s}:{k}" for _, p, s, _, k, _w in ranking[:3]]

        znak = -1 if (prop in KLUCZE_KOSZTOWE and wart and max(wart) <= 0) else 1
        propozycje.append({
            "spolka": "", "sekcja": it["sekcja"], "nazwa_oryginalna": nazwa,
            "standard_key": prop, "znak": znak, "waga": 1,
            "pewnosc": "DO_WERYFIKACJI", "zrodlo": f"kreator {spolka}".strip(),
            "uwagi": uwaga, "_pewnosc_dopasowania": pewnosc,
            "_kandydaci": " | ".join(kandydaci), "_wartosci": it["wartosci"], "_wiersz": it["wiersz"],
        })
    return propozycje, klucze_std

def nowa_spolka(plik, slownik_path="slownik.xlsx", spolka="XXX", nazwa="", arkusz=None, out=None):
    """Kreator: bada plik, proponuje konfiguracje i aliasy, zapisuje propozycja_<TICKER>.xlsx"""
    out = out or f"propozycja_{spolka}.xlsx"
    cfg = zbadaj_plik(plik, arkusz)
    cfg_do_parsera = {**cfg, "arkusz": cfg["arkusz"]}
    prop, klucze_std = zaproponuj_aliasy(plik, cfg_do_parsera, slownik_path, spolka)

    wiersz_spolki = [spolka, nazwa or spolka, os.path.basename(plik), cfg["arkusz"], cfg["kol_nazwa"],
                     cfg["kol_wartosci"], cfg["mnoznik"], cfg["naglowki_okresow"], cfg["sekcje"],
                     cfg["kapital_obrotowy_z_bilansu"], "wygenerowane przez kreator - SPRAWDZ mnoznik i sekcje"]

    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "1_Wiersz_do_Spolki"
    ws.append(["spolka", "nazwa", "plik", "arkusz", "kol_nazwa", "kol_wartosci", "mnoznik",
               "naglowki_okresow", "sekcje", "kapital_obrotowy_z_bilansu", "uwagi"])
    ws.append(wiersz_spolki)
    style_header(ws, 11)
    for col, w in zip("ABCDEFGHIJK", (10, 16, 44, 16, 12, 14, 10, 40, 70, 24, 50)):
        ws.column_dimensions[col].width = w

    ws = wb.create_sheet("2_Wykryte_sekcje")
    ws.append(["wiersz", "nazwa w pliku", "kod sekcji", "typ znacznika"])
    for s in cfg["wykryte_sekcje"]:
        ws.append([s["wiersz"], s["nazwa"], s["kod"], s["typ"]])
    style_header(ws, 4)
    for col, w in zip("ABCD", (8, 70, 14, 14)): ws.column_dimensions[col].width = w

    ws = wb.create_sheet("3_Propozycje_aliasow")
    n_kol = len(str(cfg["kol_wartosci"]).split(","))
    ws.append(["spolka", "sekcja", "nazwa_oryginalna", "standard_key", "znak", "waga", "pewnosc",
               "zrodlo", "uwagi", "| pewnosc dopasowania", "inni kandydaci", "wiersz"] +
              [f"wartosc_{i+1}" for i in range(n_kol)])
    for p in prop:
        ws.append([p["spolka"], p["sekcja"], p["nazwa_oryginalna"], p["standard_key"], p["znak"], p["waga"],
                   p["pewnosc"], p["zrodlo"], p["uwagi"], p["_pewnosc_dopasowania"], p["_kandydaci"],
                   p["_wiersz"]] + p["_wartosci"])
        kolor = {"WYSOKA": "C6EFCE", "SREDNIA": "FFEB9C", "NISKA": "FCE4D6", "BRAK": "FFC7CE"}.get(p["_pewnosc_dopasowania"], "FFFFFF")
        for c in range(1, 13):
            ws.cell(row=ws.max_row, column=c).fill = PatternFill("solid", fgColor=kolor)
    style_header(ws, 12 + n_kol)
    for col, w in zip("ABCDEFGHIJKL", (10, 12, 68, 30, 7, 7, 16, 16, 62, 20, 34, 8)):
        ws.column_dimensions[col].width = w

    ws = wb.create_sheet("4_Klucze_standardu")
    ws.append(["standard_key", "etykieta_pl"])
    for k, e in klucze_std: ws.append([k, e])
    style_header(ws, 2)
    ws.column_dimensions["A"].width = 32; ws.column_dimensions["B"].width = 52

    ws = wb.create_sheet("0_Instrukcja")
    for line in [
        [f"KREATOR NOWEJ SPOLKI - {spolka}"], [""],
        ["Krok 1. Arkusz '1_Wiersz_do_Spolki' - skopiuj ten wiersz do arkusza 'Spolki' w slownik.xlsx."],
        ["        SPRAWDZ dwie rzeczy:"],
        [f"        - mnoznik = {cfg['mnoznik']} (najwieksza liczba w pliku to {cfg['maks_wartosc']:,.0f});"],
        ["          1 = dane w tysiacach, 1000 = dane w milionach"],
        ["        - kolumna 'sekcje' - porownaj z arkuszem '2_Wykryte_sekcje'"], [""],
        ["Krok 2. Arkusz '3_Propozycje_aliasow' - przejrzyj propozycje i popraw kolumne standard_key."],
        ["        ZIELONY      = dopasowanie pewne, zwykle wystarczy zatwierdzic"],
        ["        ZOLTY        = prawdopodobne, sprawdz"],
        ["        POMARANCZOWY = slabe dopasowanie, traktuj jak podpowiedz a nie odpowiedz"],
        ["        CZERWONY     = brak podpowiedzi, zmapuj recznie (lista kluczy w arkuszu 4)"],
        ["        standard_key = POMIN oznacza swiadome pominiecie wiersza."], [""],
        ["Krok 3. Skopiuj kolumny A-I zatwierdzonych wierszy do arkusza 'Aliasy' w slownik.xlsx."],
        ["        Kolumny J, K, L (pewnosc dopasowania, kandydaci, wiersz) sluza tylko do przegladu - NIE kopiuj ich."], [""],
        ["Krok 4. Uruchom mapowanie i sprawdz arkusze 'Walidacja' oraz 'Niezmapowane':"],
        [f"        brmap.mapuj('slownik.xlsx', '{spolka}', '{os.path.basename(plik)}')"], [""],
        ["Krok 5. Powtarzaj krok 2-4 az 'Niezmapowane' bedzie puste, a 'Walidacja' cala na OK."], [""],
        ["UWAGA: kreator niczego nie dopisuje do slownika sam. Kazdy alias wymaga Twojej akceptacji."],
    ]:
        ws.append(line)
    ws.column_dimensions["A"].width = 120
    ws["A1"].font = Font(bold=True, size=14)
    wb.move_sheet("0_Instrukcja", offset=-5)
    wb.save(out)

    print(f"Wykryto: arkusz '{cfg['arkusz']}', kolumna nazw {cfg['kol_nazwa']}, kolumny wartosci {cfg['kol_wartosci']}")
    print(f"Jednostka: mnoznik {cfg['mnoznik']} (max wartosc w pliku {cfg['maks_wartosc']:,.0f}) - SPRAWDZ")
    print(f"Sekcje: {len(cfg['wykryte_sekcje'])} znacznikow "
          f"({sum(1 for s in cfg['wykryte_sekcje'] if s['typ']=='poczatek')} otwierajacych, "
          f"{sum(1 for s in cfg['wykryte_sekcje'] if s['typ']=='koniec')} zamykajacych)")
    bez = sum(1 for p in prop if p["_pewnosc_dopasowania"] == "BRAK")
    print(f"Pozycji spoza slownika: {len(prop)}  (w tym {bez} bez podpowiedzi)")
    print(f"Zapisano: {out}")
    return out

# =====================================================================
#  FUNKCJE WYSOKIEGO POZIOMU
# =====================================================================
def wczytaj_okres_wzorca(wzorzec, data=None, arkusz="Spr Fin", naglowek_okresu=None):
    """Wyciaga kolumne z pliku wzorcowego - do policzenia kapitalu obrotowego, gdy
    sprawozdanie spolki nie zawiera bilansu okresu odniesienia (uklad UoR, gdzie kolumna
    porownawcza to ten sam kwartal rok wczesniej; albo raport roczny, ktory PRZEKSZTALCA
    dane porownawcze - wtedy standard trzyma pierwotnie raportowany bilans).

    Okres wskazuje sie data bilansowa albo naglowkiem kolumny."""
    rows = list(openpyxl.load_workbook(wzorzec, read_only=True, data_only=True)[arkusz]
                .iter_rows(min_row=1, max_row=110, values_only=True))
    if naglowek_okresu:
        kol = next((i for i in range(len(rows[0]))
                    if str(rows[0][i]).strip() == str(naglowek_okresu).strip()), None)
    else:
        cel = datetime.datetime.strptime(data, "%Y-%m-%d")
        daty = next((r for r in rows if r[0] and str(r[0]).strip() == "balance_date"), None)
        kol = None if daty is None else next(
            (i for i in range(len(daty))
             if isinstance(daty[i], datetime.datetime) and daty[i] == cel), None)
    if kol is None:
        raise ValueError(f"We wzorcu '{wzorzec}' nie ma kolumny {naglowek_okresu or data}")
    return {str(r[0]).strip(): r[kol] for r in rows if r[0] and isinstance(r[0], str)}

RZYMSKIE = {"i": 1, "ii": 2, "iii": 3, "iv": 4}

def tag_okresu(naglowek, data=None):
    """Skrot okresu do nazwy pliku: "I kw 2026 / 31.03.2026" -> "1Q2026".

    Rozpoznaje kwartaly, polrocza i "N m-cy". Gdy nie ma pewnosci - zwraca sama date
    bilansowa. Naglowek z PRZESUNIETYM rokiem obrotowym ("I kw obrotowy 2025/26") tez
    idzie na date, bo "1Q2025" myloby sie z kwartalem kalendarzowym.
    """
    n = norm(naglowek)
    rok = data.year if data else None
    if rok and "obrotow" not in n:
        m = re.match(r"(i{1,3}|iv)\s+kw\b", n)
        if m:
            return f"{RZYMSKIE[m.group(1)]}Q{rok}"
        m = re.match(r"(i{1,3}|iv)\s+(polrocze|pol)\b", n)
        if m:
            return f"{RZYMSKIE[m.group(1)]}H{rok}"
        m = re.match(r"(\d+)\s*m\s*cy\b", n)
        if m:
            return f"{m.group(1)}M{rok}"
    return data.strftime("%Y-%m-%d") if data else "okres"

def mapuj(slownik_path, spolka, plik, out=None, cicho=False, bilans_poprzedni=None,
          okresy=None):
    sl = load_slownik(slownik_path)
    cfg = next((s for s in sl["spolki"] if str(s.get("spolka")).strip().upper() == spolka.upper()), None)
    if cfg is None:
        raise ValueError(f"Brak konfiguracji spolki {spolka} w arkuszu Spolki. "
                         f"Uruchom najpierw brmap.nowa_spolka(...)")
    items = parse_company(plik, cfg)
    n_kol = len(str(cfg["kol_wartosci"]).split(","))
    # Etykiety okresow mozna nadpisac z zewnatrz. Ten sam podmiot sklada raporty za rozne
    # okresy w IDENTYCZNYM ukladzie, a kod spolki jest jednoczesnie kluczem wyjatkow -
    # zalozenie drugiego kodu tylko dla podpisow kolumn gubiloby wszystkie wyjatki spolki.
    zrodlo_naglowkow = "|".join(okresy) if okresy else str(cfg.get("naglowki_okresow") or "")
    naglowki = [h.strip() for h in zrodlo_naglowkow.split("|")][:n_kol]
    while len(naglowki) < n_kol:
        naglowki.append(f"okres_{len(naglowki)+1}")
    daty = daty_z_naglowkow(zrodlo_naglowkow)
    # okres w nazwie pliku - inaczej kolejny raport tej samej spolki nadpisze poprzedni
    out = out or f"wynik_{spolka.upper()}_{tag_okresu(naglowki[0], daty[0] if daty else None)}.xlsx"

    # wariant ukladu (np. III kwartal) dziedziczy wyjatki spolki macierzystej
    kod_wyjatkow = str(cfg.get("wyjatki_z") or "").strip() or spolka
    idx = build_index(sl["aliasy"], kod_wyjatkow)
    wyniki, uzyte, audyt, niezmapowane = apply_mapping(items, idx, n_kol)
    wyniki = apply_reguly(wyniki, sl["reguly"], n_kol)
    if (str(cfg.get("kapital_obrotowy_z_bilansu") or "").strip().upper() in ("TAK", "T", "1")
            or bilans_poprzedni):
        wyniki = kapital_obrotowy_z_bilansu(wyniki, n_kol, bilans_poprzedni)
    # tolerancja sum kontrolnych: z konfiguracji spolki, a jesli brak - z jednostki
    tol = float(cfg.get("tolerancja") or 0) or (2.0 if float(cfg.get("mnoznik") or 1) < 1 else 1.0)
    # metadane, ktore da sie wyprowadzic z konfiguracji - standard je wypelnia
    # kolumna 'waluta' jest opcjonalna; przyjmujemy tylko poprawny kod (np. EUR), inaczej PLN
    waluta = str(cfg.get("waluta") or "").strip().upper()
    if not re.fullmatch(r"[A-Z]{3}", waluta):
        waluta = "PLN"
    wyniki["currency_id"] = [waluta] * n_kol
    wyniki["balance_date"] = [(daty[j].strftime("%d.%m.%Y") if j < len(daty) and daty[j] else None)
                              for j in range(n_kol)]

    szablon = str(cfg.get("szablon") or "RAP").strip().upper()
    if szablon == "RAP_B" and sl.get("klucze_b"):
        klucze, walid = sl["klucze_b"], waliduj_bank(wyniki, n_kol, tol)
    else:
        klucze, walid = sl["klucze"], waliduj(wyniki, n_kol, tol)
    zapisz(out, klucze, wyniki, naglowki, audyt, niezmapowane, walid)

    if not cicho:
        znane = [(k[1], v[0].get("standard_key")) for k, v in idx.items()]
        print(f"Zmapowano pozycji: {len(audyt)}   Niezmapowanych: {len(niezmapowane)}")
        for it in niezmapowane:
            best = sorted(((similar(it["nazwa_norm"], n), n, sk) for n, sk in znane), reverse=True)[:1]
            hint = f"   ~ moze: {best[0][2]} ({best[0][0]:.2f})" if best and best[0][0] > 0.45 else ""
            print(f"   [{it['sekcja']}] {it['nazwa'][:70]}{hint}")
        bledy = [r for r in walid if any(str(c).startswith(("ROZNICA", "PODEJRZANA")) for c in r[1:])]
        print(f"\nWalidacja: {len(walid)-len(bledy)}/{len(walid)} OK")
        for r in bledy:
            print("   !", r)
        print(f"Zapisano: {out}")
    return out, wyniki, walid, niezmapowane

def porownaj(wzorzec, data, wynik, out=None, arkusz="Spr Fin", kolumna=4, cicho=False, tol=1.0,
             naglowek_okresu=None):
    """Porownuje wynik mapowania z plikiem wzorcowym w standardzie.

    Kolumne wskazuje sie data bilansowa, a jesli ta jest we wzorcu bledna -
    naglowkiem okresu (np. "II 2026").
    """
    out = out or wynik.replace("wynik_", "porownanie_")
    wb = openpyxl.load_workbook(wzorzec, read_only=True, data_only=True)
    rows = list(wb[arkusz].iter_rows(min_row=1, max_row=110, values_only=True))
    if naglowek_okresu:
        kol = next(i for i in range(len(rows[0]))
                   if str(rows[0][i]).strip() == str(naglowek_okresu).strip())
    else:
        cel = datetime.datetime.strptime(data, "%Y-%m-%d")
        daty = next(r for r in rows if r[0] and str(r[0]).strip() == "balance_date")
        kol = next(i for i in range(len(daty)) if isinstance(daty[i], datetime.datetime) and daty[i] == cel)
    wzor = {str(r[0]).strip(): r[kol] for r in rows
            if r[0] and isinstance(r[0], str) and str(r[0]).strip() not in ("RAP", "end")}

    ws2 = openpyxl.load_workbook(wynik, data_only=True)["Wynik"]
    dane = [(r[0], r[1], r[kolumna - 1]) for r in ws2.iter_rows(min_row=2, values_only=True) if r[0]]

    out_wb = openpyxl.Workbook(); ws = out_wb.active; ws.title = "Porownanie"
    ws.append(["standard_key", "etykieta_pl", "wzorzec", "moje mapowanie", "roznica", "status"])
    zgodne = rozne = brak = 0
    for k, ev, m in dane:
        w = wzor.get(k)
        wn = w if isinstance(w, (int, float)) else None
        mn = m if isinstance(m, (int, float)) else None
        if (wn or 0) == 0 and (mn or 0) == 0:
            status, d = "obie puste", None
        elif wn is None or mn is None:
            status, d = "TYLKO W JEDNYM", None; brak += 1
        else:
            d = mn - wn
            status = "ZGODNE" if abs(d) <= tol else "ROZNICA"
            zgodne, rozne = (zgodne + 1, rozne) if status == "ZGODNE" else (zgodne, rozne + 1)
        ws.append([k, ev, wn, mn, d, status])

    FILL = {"ZGODNE": "C6EFCE", "ROZNICA": "FFC7CE", "TYLKO W JEDNYM": "FFEB9C", "obie puste": "F2F2F2"}
    for r in range(2, ws.max_row + 1):
        f = PatternFill("solid", fgColor=FILL[ws.cell(row=r, column=6).value])
        for c in range(1, 7): ws.cell(row=r, column=c).fill = f
        for c in (3, 4, 5): ws.cell(row=r, column=c).number_format = "#,##0"
    style_header(ws, 6)
    for col, w in zip("ABCDEF", (30, 48, 16, 16, 14, 18)): ws.column_dimensions[col].width = w
    out_wb.save(out)

    if not cicho:
        print(f"ZGODNE: {zgodne}   ROZNICE: {rozne}   tylko w jednym: {brak}")
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r[5] in ("ROZNICA", "TYLKO W JEDNYM"):
                print(f"   {r[0]:30} wzorzec {str(r[2]):>14}   moje {str(r[3]):>14}")
        print(f"Zapisano: {out}")
    return out, zgodne, rozne

# =====================================================================
#  COLAB
# =====================================================================
def w_colabie():
    try:
        import google.colab  # noqa
        return True
    except ImportError:
        return False

def podlacz_dysk(folder="/content/drive/MyDrive/biznesradar"):
    """Montuje Dysk Google i przechodzi do folderu roboczego (slownik zapisuje sie trwale)."""
    from google.colab import drive
    drive.mount("/content/drive")
    os.makedirs(folder, exist_ok=True)
    os.chdir(folder)
    print("Folder roboczy:", os.getcwd())
    print("Pliki:", sorted(os.listdir(".")) or "(pusty - wgraj slownik.xlsx)")
    return folder

def wgraj():
    """Okienko wyboru plikow z dysku lokalnego (tryb bez Dysku Google)."""
    from google.colab import files
    w = files.upload()
    print("Wgrano:", list(w))
    return list(w)

def pobierz(*pliki):
    from google.colab import files
    for p in pliki:
        files.download(p)

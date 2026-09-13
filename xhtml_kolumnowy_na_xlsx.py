# -*- coding: utf-8 -*-
"""XHTML z pdf2htmlEX, w ktorym KAZDA KOMORKA to osobna linia -> xlsx w ukladzie brmap.

    python xhtml_kolumnowy_na_xlsx.py raport.xhtml wynik_sprawozdania.xlsx

Sa dwa warianty wyjscia pdf2htmlEX i trzeba je rozroznic:

  * CD Projekt - caly wiersz tabeli siedzi w jednej linii, kolumny rozdzielone
    szerokimi odstepami. Do tego sluzy xhtml_na_xlsx.py.
  * Wawel - kazda komorka to OSOBNA linia: najpierw etykieta, potem (opcjonalnie)
    numer noty, potem kolejno wartosci. Tym zajmuje sie ten plik.

Rozpoznanie: jesli po zamianie odstepow prawie zadna linia nie ma podwojnej spacji,
a liczby stoja w osobnych liniach - to ten drugi wariant.
"""
import re
import sys

import openpyxl

from xhtml_na_txt import linie as linie_z_xhtml

# (regex poczatku, regex konca, naglowek wyjsciowy)
SEKCJE = [
    (r"^SPRAWOZDANIE Z SYTUACJI FINANSOWEJ$", r"^SUMA KAPITAŁU WŁASNEGO I ZOBOWIĄZAŃ$",
     "Balance Sheet (Bilans / Statement of Financial Position)"),
    (r"^SPRAWOZDANIE Z CAŁKOWITYCH DOCHODÓW$", r"^Rozwodniony zysk na akcję",
     "Income Statement (Rachunek Zyskow i Strat / P&L / Statement of Comprehensive Income)"),
    (r"^SPRAWOZDANIE Z PRZEPŁYWÓW PIENIĘŻNYCH$", r"^Środki pieniężne na koniec okresu$",
     "Cash Flow Statement (Rachunek Przeplywow Pienieznych)"),
]

NOTA = re.compile(r"^\d+(?:\.\d+)+$")                 # odsylacz do noty: 3.8, 3.21.2
LICZBA = re.compile(r"^\(?-?\d{1,3}(?:[  ]\d{3})*(?:,\d+)?\)?$|^\(?-?\d+(?:,\d+)?\)?$")
# UWAGA: nie wolno tu wpisac "^\d+$" na numery stron - taka regula zjada rowniez
# male WARTOSCI (119, 49, 0), a robi to po cichu. Numery stron i tak leza poza
# zakresami sekcji, wiec nie ma czego filtrowac.
SMIECI = re.compile(r"^(Nota$|Stan na|\d{2}\.\d{2}\.\d{4}$|\d{4} r\.$|w tys\. PLN|[|]$|"
                    r"W  A  W  E  L|.{140,}$)")
# wartosci wskaznikowe podane z jednostka ("65,50 zl") to nie pozycje sprawozdania
WSKAZNIK = re.compile(r"^[-\d  ,.()]+\s*(zł|PLN|EUR|%|szt\.?)$", re.I)


def liczba(t):
    t = t.strip().replace(" ", " ").replace(" ", " ")
    uj = t.startswith("(") and t.endswith(")")
    t = t.strip("()").replace(" ", "").replace(",", ".")
    try:
        v = float(t)
    except ValueError:
        return None
    return -v if uj else v


def wiersze(linie, ile_kolumn=2):
    """Etykieta + nastepujace po niej linie liczbowe = jeden wiersz sprawozdania."""
    out, etykieta, wartosci = [], None, []

    def zamknij():
        if etykieta:
            out.append([etykieta] + (wartosci + [None] * ile_kolumn)[:ile_kolumn])

    for s in linie:
        s = s.strip()
        if not s or SMIECI.match(s) or WSKAZNIK.match(s):
            continue
        if NOTA.match(s) and etykieta and not wartosci:
            continue                                   # odsylacz do noty miedzy etykieta a wartoscia
        if LICZBA.match(s) and etykieta is not None:
            v = liczba(s)
            if v is not None:
                wartosci.append(v)
                continue
        zamknij()
        etykieta, wartosci = s, []
    zamknij()
    return out


def konwertuj(zrodlo, cel, ile_kolumn=2):
    L = linie_z_xhtml(zrodlo)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sprawozdania"
    lacznie = 0
    for start, koniec, naglowek in SEKCJE:
        rs, rk = re.compile(start), re.compile(koniec)
        i = next((k for k, s in enumerate(L) if rs.match(s.strip())), None)
        if i is None:
            print("   nie znalazlem sekcji:", start)
            continue
        j = next((k for k in range(i + 1, len(L)) if rk.match(L[k].strip())), len(L))
        # linia konczaca tez nalezy do sekcji razem ze swoimi wartosciami
        j = min(j + 1 + ile_kolumn, len(L))
        w = wiersze(L[i + 1:j], ile_kolumn)
        ws.append([naglowek])
        ws.append(["Pozycja", "2025", "2024"][:1 + ile_kolumn])
        for r in w:
            ws.append(r)
        ws.append([])
        ws.append([])
        lacznie += len(w)
        print("   %-52s %3d pozycji" % (naglowek[:50], len(w)))
    wb.save(cel)
    print("zapisano %s (%d pozycji)" % (cel, lacznie))


if __name__ == "__main__":
    konwertuj(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "wynik_sprawozdania.xlsx")

# -*- coding: utf-8 -*-
"""Orange Polska PSr H1 2026: PDF -> xlsx w ukladzie brmap.

KOTWICA: czytamy wylacznie bloki oznaczone jako SKONSOLIDOWANE (linie 349-600).
Sprawozdanie JEDNOSTKOWE (od linii ~1608) ma identyczny uklad i jest pomijane.
Separator tysiecy w tym raporcie to KROPKA ("3.500" = 3 500), dziesietny - przecinek.
"""
import re, openpyxl

LINIE = open('/tmp/opl.txt', encoding='utf-8').read().split('\n')
TOKEN = re.compile(r'^\(?-?\d{1,3}(\.\d{3})*(,\d+)?\)?$|^-$')

def licz(t):
    t = t.strip()
    if t in ('-', ''):
        return None
    ujemna = t.startswith('(')
    t = t.strip('()').replace('.', '').replace(',', '.').replace(' ', '')
    try:
        v = float(t)
    except ValueError:
        return None
    return -v if ujemna else v

def czytaj(od, do, ile):
    """Ostatnie `ile` tokenow w linii to wartosci; reszta (bez kolumny Nota) to etykieta."""
    wynik, przed, po = [], [], False
    for linia in LINIE[od-1:do]:
        if not linia.strip():
            po = False
            continue
        tok = [t.strip() for t in re.split(r'\s{2,}', linia.strip()) if t.strip()]
        wart = [t for t in tok[-ile:] if TOKEN.match(t)]
        if len(wart) != ile:
            etyk = ' '.join(tok)
            if po and wynik and etyk[:1].islower():
                wynik[-1][0] = (wynik[-1][0] + ' ' + etyk).strip()
            else:
                przed.append(etyk)
            po = False
            continue
        etyk = ' '.join(tok[:-ile])
        # kolumna "Nota" to ostatni token etykiety, jesli wyglada jak odsylacz
        czesci = etyk.split()
        if czesci and re.fullmatch(r'\d+([.,]\d+)*', czesci[-1]):
            czesci = czesci[:-1]
        etyk = ' '.join(czesci)
        nazwa = ' '.join([p for p in przed if p] + ([etyk] if etyk else [])).strip()
        przed = []; po = True
        wynik.append([nazwa] + [licz(v) for v in wart])
    return wynik

wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Sprawozdania'
def nag(t): ws.append([t])
def blok(w):
    for r in w: ws.append(r)

nag('Income Statement (Rachunek Zyskow i Strat / P&L / Statement of Comprehensive Income)')
ws.append(['Pozycja', '3 m-ce 2026', 'I polrocze 2026 / 30.06.2026', '3 m-ce 2025', 'I polrocze 2025'])
blok(czytaj(354, 387, 4))
nag('SKONSOLIDOWANE SPRAWOZDANIE Z CALKOWITYCH DOCHODOW')
blok(czytaj(396, 411, 4))

nag('Balance Sheet (Bilans / Statement of Financial Position)')
ws.append(['Pozycja', '30.06.2026', '31.12.2025'])
blok(czytaj(425, 436, 2))     # aktywa trwale (suma zamyka blok)
blok(czytaj(437, 445, 2))     # aktywa obrotowe
blok(czytaj(447, 447, 2))     # SUMA AKTYWOW
blok(czytaj(450, 456, 2))     # kapital wlasny
blok(czytaj(458, 467, 2))     # zobowiazania dlugoterminowe
blok(czytaj(469, 479, 2))     # zobowiazania krotkoterminowe
blok(czytaj(481, 481, 2))     # SUMA PASYWOW

nag('Cash Flow Statement (Rachunek Przeplywow Pienieznych)')
ws.append(['Pozycja', '3 m-ce 2026', 'I polrocze 2026', '3 m-ce 2025', 'I polrocze 2025'])
nag('DZIALALNOSC OPERACYJNA');   blok(czytaj(544, 570, 4))
nag('DZIALALNOSC INWESTYCYJNA'); blok(czytaj(572, 589, 4))
nag('DZIALALNOSC FINANSOWA');    blok(czytaj(591, 599, 4))

wb.save('/tmp/OPL_H1_2026_sprawozdania.xlsx')
for r in ws.iter_rows(values_only=True): print(r)

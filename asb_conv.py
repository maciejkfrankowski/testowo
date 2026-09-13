# -*- coding: utf-8 -*-
"""ASBISc Enterprises, skrocone SSF H1 2026: PDF -> xlsx w ukladzie brmap.

Dane w tys. USD, format ANGIELSKI: separator tysiecy to PRZECINEK ("2,965,176"),
dziesietny to KROPKA ("1.46"). Przeliczenie na PLN robi mnoznik w arkuszu Spolki
(kurs NBP USD/PLN na dzien bilansowy) - biznesradar stosuje TEN SAM kurs
rowniez do rachunku wynikow i przeplywow, nie kurs sredni.
"""
import re, openpyxl

LINIE = open('/tmp/asb2.txt', encoding='utf-8').read().split('\n')
TOKEN = re.compile(r'^\(?-?\d{1,3}(,\d{3})*(\.\d+)?\)?$|^-$')

def licz(t):
    t = t.strip()
    if t in ('-', ''):
        return None
    ujemna = t.startswith('(')
    t = t.strip('()').replace(',', '').replace(' ', '')
    try:
        v = float(t)
    except ValueError:
        return None
    return -v if ujemna else v

def czytaj(od, do, ile=2, zastepcza=None):
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
        etyk = ' '.join(tok[:-ile]).split()
        if etyk and re.fullmatch(r'\d+([.,]\d+)*', etyk[-1]):     # kolumna "Nota" (np. "4,23")
            etyk = etyk[:-1]
        etyk = ' '.join(etyk)
        nazwa = ' '.join([p for p in przed if p] + ([etyk] if etyk else [])).strip()
        przed = []; po = True
        wynik.append([nazwa or zastepcza or '(bez nazwy)'] + [licz(v) for v in wart])
    return wynik

wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Sprawozdania'
def nag(t): ws.append([t])
def blok(w):
    for r in w: ws.append(r)

nag('Income Statement (Rachunek Zyskow i Strat / P&L / Statement of Comprehensive Income)')
ws.append(['Pozycja', 'I polrocze 2026 / 30.06.2026 (tys. USD)', 'I polrocze 2025 (tys. USD)'])
blok(czytaj(182, 201))
nag('Rozbicie wyniku i inne calkowite dochody')
blok(czytaj(202, 216, 2, 'Zysk za okres (powtorzenie)'))

nag('Balance Sheet (Bilans / Statement of Financial Position)')
ws.append(['Pozycja', '30.06.2026 (tys. USD)', '31.12.2025 (tys. USD)'])
nag('Aktywa trwale');                blok(czytaj(234, 242))
nag('Aktywa obrotowe');              blok(czytaj(245, 252))
nag('Kapital wlasny');               blok(czytaj(257, 264))
nag('Zobowiazania dlugoterminowe');  blok(czytaj(267, 270))
nag('Zobowiazania krotkoterminowe'); blok(czytaj(273, 281))

nag('Cash Flow Statement (Rachunek Przeplywow Pienieznych)')
ws.append(['Pozycja', 'I polrocze 2026 (tys. USD)', 'I polrocze 2025 (tys. USD)'])
nag('Przeplywy z dzialalnosci operacyjnej')
blok(czytaj(345, 369))
# dwie sumy operacyjne maja etykiety roznniace sie TYLKO polozeniem nawiasu
# ("(wydatki) netto" vs "(wydatki netto)") - po normalizacji sa identyczne,
# wiec pierwsza z nich dostaje tu jednoznaczna nazwe
posrednia = czytaj(370, 370)
posrednia[0][0] = 'Wplywy netto z dzialalnosci operacyjnej PRZED odsetkami i podatkiem'
blok(posrednia)
blok(czytaj(371, 373))
nag('Przeplywy z dzialalnosci inwestycyjnej');  blok(czytaj(376, 386))
nag('Przeplywy z dzialalnosci finansowej');     blok(czytaj(389, 401))

wb.save('/tmp/ASBIS_H1_2026_sprawozdania.xlsx')
for r in ws.iter_rows(values_only=True): print(r)

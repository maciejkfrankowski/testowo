# -*- coding: utf-8 -*-
"""The Farm 51 Group, raport kwartalny za II kw. 2026: PDF -> xlsx w ukladzie brmap.

USTAWA O RACHUNKOWOSCI, wariant PORONAWCZY, dane w ZLOTYCH (mnoznik 0,001).
Litery i cyfry pozycji (A., B., I., II., 1.) stoja w OSOBNEJ kolumnie PDF.
Czesc wierszy przeplywow ma tylko 2 wartosci zamiast 4 (zerowe kolumny 2026
sa w PDF pominiete) - takie wiersze dopelniamy pustymi od LEWEJ.
"""
import re, openpyxl

LINIE = open('/tmp/f51.txt', encoding='utf-8').read().split('\n')
TOKEN = re.compile(r'^-?\d{1,3}( \d{3})*,\d{2}$|^-?\d+,\d{2}$')
SMIEC = re.compile(r'Raport kwartalny za|^\s*\d{1,2}\s*$|ZYSK 128 395,74|^\s*894 265,60\s*$')

def licz(t):
    t = t.strip()
    if not t:
        return None
    try:
        return float(t.replace(' ', '').replace(' ', '').replace(',', '.'))
    except ValueError:
        return None

def czytaj(od, do, ile):
    wynik, przed, po = [], [], False
    for linia in LINIE[od-1:do]:
        if not linia.strip() or SMIEC.search(linia):
            przed.clear(); po = False
            continue
        tok = [t.strip() for t in re.split(r'\s{2,}', linia.strip()) if t.strip()]
        wart, etyk = [], []
        for t in tok:
            (wart if TOKEN.match(t) else etyk).append(t)
        if not wart:
            etyk_s = ' '.join(tok)
            if po and wynik and etyk_s[:1].islower():
                wynik[-1][0] = (wynik[-1][0] + ' ' + etyk_s).strip()
            else:
                przed.append(etyk_s)
                po = False
            continue
        nazwa = ' '.join([p for p in przed if p] + etyk).strip()
        przed = []; po = True
        wart = [None] * (ile - len(wart)) + [licz(v) for v in wart[-ile:]]
        wynik.append([nazwa or '(bez nazwy)'] + wart)
    return wynik

wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Sprawozdania'
def nag(t): ws.append([t])
def blok(w):
    for r in w: ws.append(r)

nag('Income Statement (Rachunek Zyskow i Strat / P&L / Statement of Comprehensive Income)')
ws.append(['Pozycja', 'II kw 2026', 'I polrocze 2026 / 30.06.2026', 'II kw 2025', 'I polrocze 2025'])
blok(czytaj(429, 463, 4))

nag('Balance Sheet (Bilans / Statement of Financial Position)')
ws.append(['Pozycja', '30.06.2026', '30.06.2025'])
nag('Aktywa');  blok(czytaj(382, 396, 2))
nag('Pasywa');  blok(czytaj(401, 417, 2))

nag('Cash Flow Statement (Rachunek Przeplywow Pienieznych)')
ws.append(['Pozycja', 'II kw 2026', 'I polrocze 2026', 'II kw 2025', 'I polrocze 2025'])
nag('Przeplywy z dzialalnosci operacyjnej');    blok(czytaj(480, 494, 4))
nag('Przeplywy z dzialalnosci inwestycyjnej');  blok(czytaj(497, 503, 4))
nag('Przeplywy z dzialalnosci finansowej');     blok(czytaj(507, 524, 4))

wb.save('/tmp/FARM51_H1_2026_sprawozdania.xlsx')
for r in ws.iter_rows(values_only=True): print(r)

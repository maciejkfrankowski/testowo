# -*- coding: utf-8 -*-
"""Grupa PZU, skrocone SSF za I polrocze 2026: PDF -> xlsx w ukladzie brmap.

Dane w MLN PLN. Bilans jest PLASKI - bez podzialu na aktywa trwale/obrotowe
i bez zobowiazan dlugoterminowych (wszystko jest krotkoterminowe). Caly blok
aktywow trafia do jednej sekcji AKTYWA, caly blok zobowiazan do ZOB_KR;
sumy aktywow trwalych i obrotowych wyliczaja reguly obliczane.
RZiS ma 4 kolumny (kwartal / narastajaco), bilans i przeplywy po 2.
"""
import re, openpyxl

LINIE = open('/tmp/pzu.txt', encoding='utf-8').read().split('\n')
TOKEN = re.compile(r'^\(?-?\d{1,3}( \d{3})*(,\d+)?\)?$|^-$')

def licz(t):
    t = t.strip()
    if t in ('-', ''):
        return None
    ujemna = t.startswith('(')
    t = t.strip('()').replace(' ', '').replace(' ', '').replace(',', '.')
    try:
        v = float(t)
    except ValueError:
        return None
    return -v if ujemna else v

def czytaj(od, do, ile):
    wynik, przed, po = [], [], False
    SMIEC = re.compile(r'Grupa Kapitałowa Powszechnego|Skrócone śródroczne skonsolidowane sprawozdanie finansowe|w milionach złotych|Śródroczne skonsolidowane sprawozdanie z przep|\(kontynuacja\)|1 stycznia –|Skonsolidowane sprawozdanie z przepływów|^\s*Nota\s*$|30 czerwca 20')
    for linia in LINIE[od-1:do]:
        if SMIEC.search(linia):
            przed.clear(); po = False       # stopka strony przerywa sklejanie etykiety
            continue
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
        if etyk and re.fullmatch(r'\d+(\.\d+)*', etyk[-1]):     # kolumna "Nota"
            etyk = etyk[:-1]
        etyk = ' '.join(etyk)
        nazwa = ' '.join([p for p in przed if p] + ([etyk] if etyk else [])).strip()
        nazwa = re.sub(r'^\d+(\.\d+)+\s+', '', nazwa)      # numer noty z osobnej linii NAD wierszem
        przed = []; po = True
        wynik.append([nazwa or '(bez nazwy)'] + [licz(v) for v in wart])
    return wynik

wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Sprawozdania'
def nag(t): ws.append([t])
def blok(w):
    for r in w: ws.append(r)

nag('Income Statement (Rachunek Zyskow i Strat / P&L / Statement of Comprehensive Income)')
ws.append(['Pozycja', 'II kw 2026', 'I polrocze 2026 / 30.06.2026', 'II kw 2025', 'I polrocze 2025'])
blok(czytaj(62, 115, 4))

nag('Balance Sheet (Bilans / Statement of Financial Position)')
ws.append(['Pozycja', '30.06.2026', '31.12.2025'])
nag('Aktywa (uklad plaski)');        blok(czytaj(189, 211, 2))
nag('Kapital wlasny');               blok(czytaj(226, 237, 2))
nag('Zobowiazania (uklad plaski)');  blok(czytaj(240, 254, 2))
nag('Pasywa razem');                 blok(czytaj(255, 255, 2))

nag('Cash Flow Statement (Rachunek Przeplywow Pienieznych)')
ws.append(['Pozycja', 'I polrocze 2026', 'I polrocze 2025'])
nag('Przeplywy z dzialalnosci operacyjnej');    blok(czytaj(368, 395, 2))
nag('Przeplywy z dzialalnosci inwestycyjnej');  blok(czytaj(398, 439, 2))
nag('Przeplywy z dzialalnosci finansowej');     blok(czytaj(441, 458, 2))

wb.save('/tmp/PZU_H1_2026_sprawozdania.xlsx')
for r in ws.iter_rows(values_only=True): print(r)

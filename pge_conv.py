# -*- coding: utf-8 -*-
"""GK PGE, skonsolidowane SF za 2025: PDF -> xlsx w ukladzie brmap.
Dane w MLN PLN, separator tysiecy to KROPKA ("61.434" = 61 434 mln).
Bilans ma 3 kolumny dat, RZiS i przeplywy po 2.
"""
import re, openpyxl

LINIE = open('/tmp/pge.txt', encoding='utf-8').read().split('\n')
TOKEN = re.compile(r'^\(?-?\d{1,3}(\.\d{3})*(,\d+)?\)?$|^-$')
BEZ_NAZWY = {'ZOB_KR': 'RAZEM ZOBOWIĄZANIA'}

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

def czytaj(od, do, ile, zastepcza=None):
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
        if etyk and re.fullmatch(r'\d+(\.\d+)*', etyk[-1]):     # kolumna "Nota"
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
ws.append(['Pozycja', 'IV kw 2025 / 31.12.2025', 'IV kw 2024 / 31.12.2024'])
blok(czytaj(269, 313, 2))

nag('Balance Sheet (Bilans / Statement of Financial Position)')
ws.append(['Pozycja', '31.12.2025', '31.12.2024', '01.01.2024'])
nag('Aktywa trwale');                blok(czytaj(335, 346, 3))
nag('Aktywa obrotowe');              blok(czytaj(348, 357, 3))
blok(czytaj(358, 358, 3))
nag('Kapital wlasny');               blok(czytaj(360, 370, 3))
nag('Zobowiazania dlugoterminowe');  blok(czytaj(372, 379, 3))
nag('Zobowiazania krotkoterminowe'); blok(czytaj(381, 390, 3, 'RAZEM ZOBOWIĄZANIA'))
blok(czytaj(391, 391, 3))

nag('Cash Flow Statement (Rachunek Przeplywow Pienieznych)')
ws.append(['Pozycja', 'IV kw 2025', 'IV kw 2024'])
nag('Przeplywy z dzialalnosci operacyjnej');    blok(czytaj(507, 522, 2))
nag('Przeplywy z dzialalnosci inwestycyjnej');  blok(czytaj(525, 534, 2))
nag('Przeplywy z dzialalnosci finansowej');     blok(czytaj(537, 548, 2))

wb.save('/tmp/PGE_2025_sprawozdania.xlsx')
for r in ws.iter_rows(values_only=True): print(r)

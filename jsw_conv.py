# -*- coding: utf-8 -*-
"""JSW Q1 2026: PDF (pdftotext -layout) -> xlsx w ukladzie brmap."""
import re, openpyxl

LINIE = open('/tmp/jsw.txt', encoding='utf-8').read().split('\n')

WART  = re.compile(r'^\(?-?\d[\d  ]*,\d+\)?$')
NOTA  = re.compile(r'^\d+(\.\d+)*[a-z]?(,\s*\d+(\.\d+)*)*$')
SMIEC = re.compile(r'stanowią jego integralną|ŚRÓDROCZNE SKRÓCONE|GRUPY KAPITAŁOWEJ|ZA OKRES 3 MIESIĘCY'
                   r'|Wszystkie kwoty w tabelach|dane przekształcone|^\s*\d{1,3}\s*$|Szczegóły dotyczące przekształcenia')

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

def rozbij(linia):
    etyk, wart = [], []
    for t in re.split(r'\s{2,}', linia.strip()):
        t = t.strip()
        if not t:
            continue
        if WART.match(t) or t == '-':
            wart.append(t)
        elif NOTA.match(t):
            continue                       # kolumna "Nota"
        else:
            etyk.append(t)
    return ' '.join(etyk).strip(), wart

def czytaj(od, do):
    """Wiersze z zakresu; sklejanie etykiet lamanych przez konwerter PDF."""
    wynik, przed, po_danych = [], [], False
    for linia in LINIE[od-1:do]:
        if not linia.strip() or SMIEC.search(linia):
            po_danych = False          # pusta linia konczy wiersz
            continue
        etyk, wart = rozbij(linia)
        if not wart:
            # fragment etykiety lamanej PO liczbach: sasiaduje z wierszem i zaczyna sie mala litera
            if po_danych and wynik and etyk[:1].islower() and not etyk.lower().startswith('w tym'):
                wynik[-1][0] = (wynik[-1][0] + ' ' + etyk).strip()
            else:
                przed.append(etyk)
            po_danych = False
            continue
        po_danych = True
        nazwa = ' '.join([p for p in przed if p] + ([etyk] if etyk else [])).strip()
        przed = []
        wynik.append([nazwa, licz(wart[0]), licz(wart[1]) if len(wart) > 1 else None])
    if przed and wynik:                    # ogon etykiety po ostatniej liczbie
        wynik[-1][0] = (wynik[-1][0] + ' ' + ' '.join(przed)).strip()
    return wynik

wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Sprawozdania'
def naglowek(t): ws.append([t])
def wiersze(ws_rows):
    for w in ws_rows: ws.append(w)

naglowek('Income Statement (Rachunek Zyskow i Strat / P&L / Statement of Comprehensive Income)')
ws.append(['Pozycja', '3 m-ce 2026 / 31.03.2026', '3 m-ce 2025 / 31.03.2025'])
wiersze(czytaj(132, 175))

naglowek('Balance Sheet (Bilans / Statement of Financial Position)')
ws.append(['Pozycja', '31.03.2026', '31.12.2025'])
naglowek('AKTYWA TRWAŁE');                 wiersze(czytaj(197, 206))
naglowek('AKTYWA OBROTOWE');               wiersze(czytaj(209, 220))
naglowek('KAPITAŁ WŁASNY');                wiersze(czytaj(240, 248))
naglowek('ZOBOWIĄZANIA DŁUGOTERMINOWE');   wiersze(czytaj(253, 259))
naglowek('ZOBOWIĄZANIA KRÓTKOTERMINOWE');  wiersze(czytaj(263, 272))
wiersze(czytaj(275, 276))

naglowek('Cash Flow Statement (Rachunek Przeplywow Pienieznych)')
ws.append(['Pozycja', '3 m-ce 2026', '3 m-ce 2025'])
naglowek('PRZEPŁYWY PIENIĘŻNE Z DZIAŁALNOŚCI OPERACYJNEJ');    wiersze(czytaj(345, 348))
naglowek('PRZEPŁYWY PIENIĘŻNE Z DZIAŁALNOŚCI INWESTYCYJNEJ');  wiersze(czytaj(352, 363))
naglowek('PRZEPŁYWY PIENIĘŻNE Z DZIAŁALNOŚCI FINANSOWEJ');     wiersze(czytaj(367, 377))

naglowek('Noty (uzupelnienie)');                               wiersze(czytaj(2750, 2765))

wb.save('/tmp/JSW_Q1_2026_sprawozdania.xlsx')
for r in ws.iter_rows(values_only=True): print(r)

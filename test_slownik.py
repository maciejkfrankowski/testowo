# -*- coding: utf-8 -*-
"""Kontrola spojnosci slownika. NIE wymaga zadnych danych spolek.

    python test_slownik.py

Lapie bledy, ktore powstaja przy recznej edycji w Excelu: literowke w standard_key,
alias wpisany dwa razy, regule odwolujaca sie do nieistniejacego klucza, zla konfiguracje
sekcji. Kazdy z nich objawia sie pozniej jako cicha bzdura w wyniku, wiec lepiej zlapac tu.
"""
import os
import re
import sys

import brmap

TU = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
SLOWNIK = os.path.join(TU, sys.argv[1] if len(sys.argv) > 1 else "slownik.xlsx")

# klucze pomocnicze - nie trafiaja do arkusza Wynik, ale wolno sie do nich odwolywac
POMOCNICZE = {"kapitaly_pozostale_nierozdzielone", "retained_earnings_total",
              "change_in_wc_nierozdzielony"}
SPECJALNE = {"POMIN", "IGNORUJ"}

bledy, ostrzezenia = [], []


def sprawdz():
    sl = brmap.load_slownik(SLOWNIK)
    klucze = {str(k.get("standard_key")).strip() for k in sl["klucze"] if k.get("standard_key")}
    # szablon bankowy ma wlasny zestaw kluczy - oba sa poprawnymi celami
    klucze |= {str(k.get("standard_key")).strip() for k in sl.get("klucze_b", []) if k.get("standard_key")}
    dozwolone = klucze | POMOCNICZE | SPECJALNE

    # --- 1. aliasy wskazuja na istniejace klucze ---
    for i, a in enumerate(sl["aliasy"], 2):
        sk = str(a.get("standard_key") or "").strip()
        nazwa = str(a.get("nazwa_oryginalna") or "").strip()
        if not nazwa:
            bledy.append(f"Aliasy w.{i}: pusta nazwa_oryginalna")
            continue
        if not sk:
            bledy.append(f"Aliasy w.{i}: pusty standard_key dla '{nazwa[:45]}'")
        elif sk not in dozwolone:
            bledy.append(f"Aliasy w.{i}: nieznany standard_key '{sk}' dla '{nazwa[:45]}'")
        znak = a.get("znak")
        if znak is not None and float(znak) not in (1.0, -1.0):
            ostrzezenia.append(f"Aliasy w.{i}: nietypowy znak {znak} dla '{nazwa[:45]}'")
        if not str(a.get("sekcja") or "").strip():
            bledy.append(f"Aliasy w.{i}: pusta sekcja dla '{nazwa[:45]}'")

    # --- 2. brak zdublowanych regul (ten sam skutek dwa razy = podwojne liczenie) ---
    widziane = {}
    for i, a in enumerate(sl["aliasy"], 2):
        sig = (str(a.get("spolka") or "").strip().upper(),
               str(a.get("sekcja") or "").strip(),
               brmap.norm(a.get("nazwa_oryginalna")),
               str(a.get("standard_key") or "").strip(),
               float(a.get("znak") or 1), float(a.get("waga") or 1))
        if sig in widziane:
            ostrzezenia.append(f"Aliasy w.{i}: powtorzenie wiersza {widziane[sig]} "
                               f"({sig[1]} / {sig[2][:40]} -> {sig[3]})")
        widziane[sig] = i

    # --- 3. ta sama nazwa w tej samej sekcji nie moze zasilac klucza i byc pominieta ---
    wg_klucza = {}
    for a in sl["aliasy"]:
        k = (str(a.get("spolka") or "").strip().upper(),
             str(a.get("sekcja") or "").strip(), brmap.norm(a.get("nazwa_oryginalna")))
        wg_klucza.setdefault(k, set()).add(str(a.get("standard_key") or "").strip())
    for k, cele in wg_klucza.items():
        if "POMIN" in cele and len(cele) > 1:
            bledy.append(f"Sprzecznosc: [{k[1]}] '{k[2][:45]}' jest jednoczesnie POMIN i {sorted(cele - {'POMIN'})}")

    # --- 4. reguly odwoluja sie do znanych kluczy ---
    for i, r in enumerate(sl["reguly"], 2):
        sk = str(r.get("standard_key") or "").strip()
        if sk and sk not in dozwolone:
            bledy.append(f"Reguly w.{i}: nieznany standard_key '{sk}'")
        for _, nazwa in re.findall(r"([+-]?)\s*([a-z_0-9]+)", str(r.get("formula") or "")):
            if nazwa not in dozwolone:
                bledy.append(f"Reguly w.{i}: formula odwoluje sie do nieznanego '{nazwa}'")

    # --- 5. konfiguracja spolek daje sie sparsowac ---
    for i, s in enumerate(sl["spolki"], 2):
        tick = str(s.get("spolka") or "").strip()
        if not tick:
            bledy.append(f"Spolki w.{i}: pusty skrot spolki")
            continue
        try:
            [int(c) for c in str(s.get("kol_wartosci")).split(",")]
        except ValueError:
            bledy.append(f"Spolki w.{i} ({tick}): kol_wartosci '{s.get('kol_wartosci')}' to nie sa numery kolumn")
        try:
            float(s.get("mnoznik") or 1)
        except (TypeError, ValueError):
            bledy.append(f"Spolki w.{i} ({tick}): mnoznik '{s.get('mnoznik')}' nie jest liczba")
        sek = str(s.get("sekcje") or "")
        if not sek.strip():
            bledy.append(f"Spolki w.{i} ({tick}): pusta kolumna 'sekcje'")
        for para in sek.split("|"):
            if para.strip() and "=>" not in para:
                bledy.append(f"Spolki w.{i} ({tick}): znacznik bez '=>': {para[:45]}")

    print(f"Slownik: {os.path.basename(SLOWNIK)}")
    print(f"  kluczy standardu : {len(klucze)}  (w tym bankowe: {len(sl.get('klucze_b', []))})")
    print(f"  aliasow          : {len(sl['aliasy'])}")
    print(f"  regul            : {len(sl['reguly'])}")
    print(f"  spolek           : {len(sl['spolki'])}")


sprawdz()
print()
for o in ostrzezenia:
    print(f"  OSTRZEZENIE: {o}")
for b in bledy:
    print(f"  BLAD: {b}")
print()
if bledy:
    print(f"NIEPOWODZENIE - bledow: {len(bledy)}, ostrzezen: {len(ostrzezenia)}")
    sys.exit(1)
print(f"OK - slownik spojny (ostrzezen: {len(ostrzezenia)})")

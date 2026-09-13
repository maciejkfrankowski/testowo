# -*- coding: utf-8 -*-
"""
start.py - interaktywny program do mapowania sprawozdan na standard biznesradar.

Uruchom:  python start.py
Program prowadzi krok po kroku: pyta o spolke, prosi o wskazanie plikow z dysku,
a na koncu zapisuje plik Excel z wynikiem.
"""
import os
import re
import sys
import datetime

# katalog skryptu do sciezki importu; folderu roboczego NIE zmieniamy przy imporcie,
# bo w Colabie ustawia go notatnik (np. na folder Dysku Google)
_TU = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
if _TU not in sys.path:
    sys.path.insert(0, _TU)

try:
    import openpyxl  # noqa
except ImportError:
    print("Brak biblioteki openpyxl. Uruchom:")
    print(f"   {sys.executable} -m pip install openpyxl")
    sys.exit(1)
import brmap

KRESKA = "=" * 68
SLOWNIK = "slownik.xlsx"


def naglowek(txt):
    print("\n" + KRESKA)
    print(txt)
    print(KRESKA)


class BrakWejscia(Exception):
    pass


def pytaj(tekst, domyslnie=""):
    try:
        odp = input(tekst).strip()
    except (EOFError, OSError):
        # tak konczy sie "!python start.py" w Colabie - podproces nie ma stdin
        raise BrakWejscia(
            "Program nie moze zadac pytania - brak wejscia z klawiatury.\n"
            "W Colabie NIE uruchamiaj przez '!python start.py'.\n"
            "Zamiast tego w komorce napisz:\n"
            "    import start\n"
            "    start.main()")
    return odp or domyslnie


def pliki_xlsx(pomijaj=()):
    return sorted(f for f in os.listdir(".")
                  if f.lower().endswith((".xlsx", ".xlsm"))
                  and not f.startswith("~$")
                  and f not in pomijaj
                  and not f.lower().startswith("slownik")
                  and not f.startswith(("wynik_", "porownanie_", "propozycja_")))


def wybierz_plik(tekst, pomijaj=(), wymagany=True, domyslny=None):
    """Lista ponumerowana + mozliwosc wklejenia pelnej sciezki."""
    lista = pliki_xlsx(pomijaj)
    print("\n" + tekst)
    if domyslny and os.path.exists(domyslny):
        print(f"   Enter) {domyslny}   <- domyslnie")
    if lista:
        for i, f in enumerate(lista, 1):
            print(f"   {i:2}) {f}")
    else:
        print("   (brak plikow .xlsx w tym folderze)")
    print("    0) wklej pelna sciezke do pliku")
    if not wymagany and not (domyslny and os.path.exists(domyslny)):
        print("   Enter) pomin")

    while True:
        odp = pytaj("\n   Wybor: ")
        if not odp and domyslny and os.path.exists(domyslny):
            return domyslny
        if not odp and not wymagany:
            return None
        if odp == "0":
            sciezka = pytaj("   Sciezka: ").strip('"').strip("'")
            if os.path.exists(sciezka):
                return sciezka
            print("   Nie ma takiego pliku.")
            continue
        if odp.isdigit() and 1 <= int(odp) <= len(lista):
            return lista[int(odp) - 1]
        if os.path.exists(odp.strip('"')):
            return odp.strip('"')
        print("   Nieprawidlowy wybor.")


DATA_RE = re.compile(r"(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})")


def daty_z_naglowkow(naglowki):
    """Wyciaga daty bilansowe z opisow okresow (bierze ostatnia date w kazdym opisie)."""
    out = []
    for h in str(naglowki or "").split("|"):
        znalezione = DATA_RE.findall(h)
        if not znalezione:
            out.append(None)
            continue
        d, m, y = znalezione[-1]
        y = int(y)
        y += 2000 if y < 100 else 0
        try:
            out.append(datetime.datetime(y, int(m), int(d)))
        except ValueError:
            out.append(None)
    return out


def poprzedni_kwartal(data):
    m, y = data.month - 3, data.year
    if m <= 0:
        m += 12
        y -= 1
    dni = [31, 29 if y % 4 == 0 and (y % 100 != 0 or y % 400 == 0) else 28, 31, 30, 31, 30,
           31, 31, 30, 31, 30, 31][m - 1]
    return datetime.datetime(y, m, min(data.day, dni))


def wczytaj_konfiguracje(ticker):
    if not os.path.exists(SLOWNIK):
        return None, None
    sl = brmap.load_slownik(SLOWNIK)
    cfg = next((s for s in sl["spolki"]
                if str(s.get("spolka") or "").strip().upper() == ticker.upper()), None)
    return sl, cfg


def uruchom_mapowanie(ticker, plik, wzorzec, cfg, okresy=None):
    """Mapuje i - jesli podano wzorzec - porownuje. Zwraca sciezke do wyniku."""
    mnoznik = float(cfg.get("mnoznik") or 1)
    tol = 2.0 if mnoznik < 1 else 1.0

    daty = daty_z_naglowkow("|".join(okresy) if okresy else cfg.get("naglowki_okresow"))
    prev = None
    if len(daty) >= 2 and daty[0] and daty[1]:
        odstep = (daty[0] - daty[1]).days
        if odstep > 200:                       # kolumna porownawcza to rok wczesniej
            print(f"\n   UWAGA: kolumna porownawcza to {daty[1]:%d.%m.%Y}, czyli rok przed "
                  f"{daty[0]:%d.%m.%Y}.")
            print("   Kapitalu obrotowego nie da sie policzyc z samego sprawozdania -")
            print("   trzeba dociagnac bilans poprzedniego kwartalu z pliku wzorcowego.")
            if wzorzec:
                cel = poprzedni_kwartal(daty[0])
                try:
                    prev = brmap.wczytaj_okres_wzorca(wzorzec, cel.strftime("%Y-%m-%d"))
                    print(f"   OK - pobrano bilans na {cel:%d.%m.%Y} z pliku {wzorzec}.")
                except StopIteration:
                    print(f"   Nie znalazlem kolumny {cel:%d.%m.%Y} we wzorcu - "
                          f"kapital obrotowy bedzie z rachunku przeplywow.")
            else:
                print("   Bez pliku wzorcowego kapital obrotowy bedzie wziety z przeplywow spolki.")

    naglowek(f"MAPOWANIE {ticker}")
    out, wyniki, walid, niezmapowane = brmap.mapuj(SLOWNIK, ticker, plik,
                                                   bilans_poprzedni=prev, okresy=okresy)

    if wzorzec:
        naglowek("POROWNANIE ZE WZORCEM")
        data_bil = daty[0] if daty and daty[0] else None
        if data_bil:
            try:
                _, zgodne, rozne = brmap.porownaj(wzorzec, data_bil.strftime("%Y-%m-%d"), out, tol=tol)
                if zgodne == 0 and rozne > 5:
                    print("\n   !! ANI JEDNEJ ZGODNEJ POZYCJI.")
                    print(f"      To prawie na pewno nie jest wzorzec tej samej spolki")
                    print(f"      albo nie ten okres. Sprawdz plik '{wzorzec}'.")
                elif rozne > zgodne:
                    print(f"\n   !! Wiecej roznic ({rozne}) niz zgodnosci ({zgodne}) - sprawdz wzorzec i okres.")
            except StopIteration:
                print(f"   We wzorcu nie ma kolumny z data {data_bil:%d.%m.%Y} - pomijam porownanie.")
        else:
            print("   Nie udalo sie odczytac daty bilansowej - pomijam porownanie.")
    return out, niezmapowane


def main(folder=None):
    """folder - katalog z slownik.xlsx i plikami sprawozdan.
    Domyslnie: biezacy katalog (Colab) albo katalog skryptu (uruchomienie lokalne)."""
    if folder:
        os.chdir(folder)
    naglowek("MAPOWANIE SPRAWOZDAN NA STANDARD BIZNESRADAR")
    print(f"Folder roboczy: {os.getcwd()}")

    if not os.path.exists(SLOWNIK):
        print(f"\nBRAK PLIKU {SLOWNIK} w tym folderze.")
        if os.path.isdir("slownik_csv"):
            print("Sa za to CSV-ki. Odtworz z nich slownik:")
            print(f"   {sys.executable} slownik_z_csv.py")
        else:
            print("Wgraj slownik do tego samego katalogu co start.py i uruchom ponownie.")
        return
    sl = brmap.load_slownik(SLOWNIK)
    znane = [str(s.get("spolka")).strip() for s in sl["spolki"] if s.get("spolka")]
    print(f"Slownik: {SLOWNIK}  ({len(sl['aliasy'])} aliasow, spolki: {', '.join(znane)})")

    ticker = ""
    while not ticker:
        ticker = pytaj("\nPodaj skrot spolki (np. NEU, KGH, AOL): ").upper()
    nowa = ticker not in [z.upper() for z in znane]

    plik = wybierz_plik(f"Wskaz plik ze sprawozdaniem spolki {ticker}:")
    wzorzec = wybierz_plik(
        f"Wskaz plik wzorcowy dla {ticker} w standardzie "
        f"(bez niego nie ma kontroli poprawnosci):",
        pomijaj=(plik,), wymagany=False, domyslny=f"{ticker}.xlsx")

    # kontrola: latwo pomylic sie o jedna pozycje na liscie i porownac z cudzym wzorcem
    if wzorzec:
        rdzen = os.path.splitext(os.path.basename(wzorzec))[0].upper()
        if ticker.upper() not in rdzen:
            print(f"\n   !! Wybrany wzorzec to '{wzorzec}', a spolka to {ticker}.")
            print("      Nazwa pliku nie pasuje do skrotu spolki.")
            if not pytaj("      Na pewno kontynuowac? [t/N]: ", "N").lower().startswith("t"):
                wzorzec = None
                print("      Pomijam wzorzec.")

    if nowa:
        naglowek(f"NOWA SPOLKA {ticker} - KREATOR")
        nazwa = pytaj("Pelna nazwa spolki (Enter = pomin): ", ticker)
        prop = brmap.nowa_spolka(plik, SLOWNIK, ticker, nazwa)

        print("\n" + "-" * 68)
        print(f"Otworz {prop} i wykonaj dwie rzeczy:")
        print(f"  1. Arkusz '1_Wiersz_do_Spolki' -> skopiuj wiersz do arkusza 'Spolki' w {SLOWNIK}")
        print("     (sprawdz kolumne 'mnoznik': 1 = tysiace, 1000 = miliony, 0.001 = zlote)")
        print(f"  2. Arkusz '3_Propozycje_aliasow' -> popraw kolumne 'standard_key',")
        print(f"     potem skopiuj kolumny A-I do arkusza 'Aliasy' w {SLOWNIK}")
        print("     zielony = pewne, zolty = sprawdz, pomaranczowy = slabe, czerwony = zmapuj recznie")
        print("-" * 68)
        pytaj("\nGdy zapiszesz slownik, wcisnij Enter zeby zmapowac... ")

    okresy = None
    while True:
        sl, cfg = wczytaj_konfiguracje(ticker)
        if cfg is None:
            print(f"\nW arkuszu 'Spolki' nadal nie ma wiersza dla {ticker}.")
            if pytaj("Sprobowac jeszcze raz? [T/n]: ", "T").lower().startswith("n"):
                return
            continue

        if okresy is None:
            # Ta sama spolka sklada raporty za rozne okresy w identycznym ukladzie.
            # Kod spolki jest kluczem wyjatkow, wiec NIE zakladamy drugiego kodu tylko
            # po to, zeby zmienic podpisy kolumn - podajemy je tutaj.
            print(f"\nOkresy ze slownika: {cfg.get('naglowki_okresow')}")
            podane = pytaj("Inne okresy? (rozdziel '|', Enter = zostaw): ")
            okresy = [x.strip() for x in podane.split("|") if x.strip()] if podane else []

        try:
            out, niezmapowane = uruchom_mapowanie(ticker, plik, wzorzec, cfg, okresy or None)
        except BrakWejscia:
            raise
        except Exception as e:
            print(f"\nBLAD: {e}")
            if pytaj("\nPoprawic slownik i sprobowac ponownie? [T/n]: ", "T").lower().startswith("n"):
                return
            continue

        naglowek("WYNIK")
        print(f"Plik Excel: {os.path.abspath(out)}")
        print("  arkusz 'Wynik'        - kolumna gotowa do wklejenia do standardu")
        print("  arkusz 'Walidacja'    - sumy kontrolne")
        print("  arkusz 'Audyt'        - skad wzieta jest kazda liczba")
        print("  arkusz 'Niezmapowane' - co jeszcze zostalo")
        if niezmapowane:
            print(f"\nUWAGA: {len(niezmapowane)} pozycji bez mapowania - uzupelnij arkusz 'Aliasy'.")

        if pytaj("\nPoprawic slownik i przeliczyc jeszcze raz? [t/N]: ", "N").lower().startswith("t"):
            pytaj("Zapisz zmiany w slowniku i wcisnij Enter... ")
            continue
        break

    print("\nGotowe.\n")


if __name__ == "__main__":
    try:
        os.chdir(_TU)          # uruchomienie lokalne: pracuj w katalogu skryptu
        main()
    except KeyboardInterrupt:
        print("\nPrzerwano.")
    except BrakWejscia as e:
        print(f"\n{e}")

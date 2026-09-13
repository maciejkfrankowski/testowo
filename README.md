# Mapowanie sprawozdań na standard biznesradar

Zamienia plik sprawozdania spółki giełdowej na kolumnę w standardowym nazewnictwie
(`total_assets`, `revenues`, `ebit`, ...). Obsługuje MSSF i ustawę o rachunkowości.

## Szybki start

Po sklonowaniu repo najpierw odtwórz słownik z CSV:

```bash
python slownik_z_csv.py
```

Potem na Windows kliknij **`start.bat`** — sam sprawdzi Pythona i doinstaluje `openpyxl`, jeśli brakuje.

Ręcznie:

```bash
python -m pip install -r requirements.txt
python start.py
```

Program pyta o skrót spółki, prosi o wskazanie plików z dysku i zapisuje `wynik_TICKER_OKRES.xlsx`.

W Google Colab — w komórce notatnika (**nie** przez `!python start.py`, podproces nie ma klawiatury):

```python
import start
start.main()
```

## Źródło prawdy: `slownik.xlsx`

Cała wiedza o mapowaniu siedzi w Excelu, nie w kodzie. W `.py` nie ma ani jednej
polskiej nazwy pozycji — programy tylko czytają słownik. Nowa spółka to zmiana
danych, nie kodu.

Obieg pracy:

```bash
# edytujesz slownik.xlsx w Excelu (arkusze Aliasy i Spolki)
python slownik_csv.py     # przed commitem: xlsx -> CSV, żeby git pokazał różnice
git commit -am "dodane aliasy ..."

git pull
python slownik_z_csv.py   # po pobraniu: CSV -> xlsx
```

CSV-ki w `slownik_csv/` to **kopia do wersjonowania**, nie osobne źródło. Excel
wygrywa; CSV istnieje dlatego, że `git diff` nie potrafi pokazać, co zmieniło się
w pliku binarnym.

> **`build_slownik.py` nadpisuje słownik.** Bez flagi `--nadpisz` skrypt się nie
> uruchomi, a przed zapisem robi kopię. Trzymamy go w repo jako zapis decyzji —
> przy każdym aliasie jest komentarz, dlaczego akurat tak — i jako ratunek, gdyby
> plik Excela przepadł. Do codziennej pracy nie jest potrzebny.

## Pliki

| plik | rola |
|---|---|
| `slownik.xlsx` | **źródło prawdy** — aliasy, reguły, konfiguracja spółek |
| `slownik_csv/` | kopia słownika w CSV, do wersjonowania w gicie |
| `slownik_csv.py` | xlsx -> CSV (przed commitem) |
| `slownik_z_csv.py` | CSV -> xlsx (po `git pull`) |
| `build_slownik.py` | odtwarza słownik od zera + historia decyzji; wymaga `--nadpisz` |
| `brmap.py` | cała logika (parser, mapowanie, walidacja, porównanie, kreator) |
| `maper.py` | uproszczony maper jednoplikowy, pod Colab |
| `start.py` | program interaktywny |
| `map.py`, `porownaj.py` | nakładki wiersza poleceń na `brmap.py` |
| `pdf_na_xlsx.py` | konwerter: sprawozdanie w PDF -> xlsx (tryb `--dopisz` dokleja notę) |
| `xhtml_na_xlsx.py` | konwerter: XHTML z pdf2htmlEX -> xlsx |
| `xhtml_na_txt.py` | pomocniczy: XHTML -> tekst (używany przez powyższy) |
| `hist.py` | diagnostyka pliku wzorcowego: narastanie, kapitał obrotowy, pola meta |
| `biznesradar_colab.ipynb` | notatnik Colab |

## Ta sama spółka, inny okres

Kod spółki jest **jednocześnie kluczem wyjątków** w arkuszu `Aliasy`. Nie zakładaj
drugiego kodu (`OPN_H`, `OPN_2026` itp.) dla raportu za inny okres — nowy kod nie
odziedziczy żadnego wyjątku spółki i mapowanie się rozjedzie, zwykle na znakach kosztów.

Raporty kwartalne i półroczne tej samej spółki mają z reguły identyczny układ, więc
wystarczy ten sam kod. Zmieniają się tylko podpisy kolumn — program pyta o nie przy
starcie:

```
Okresy ze slownika: I kw 2026 / 31.03.2026|IV kw 2025 / 31.12.2025
Inne okresy? (rozdziel '|', Enter = zostaw): I polrocze 2026 / 30.06.2026|IV kw 2025 / 31.12.2025
```

Podane okresy trafiają do nagłówków kolumn **i do pola `balance_date`**.

Jeśli zmienia się nie tylko podpis, ale **układ kolumn** — bo raport za III kwartał
podaje osobno kwartał i narastająco, a roczny tylko rok — załóż wiersz wariantu
z kolumną `wyjatki_z`:

| spolka | kol_wartosci | kol_wartosci_bilans | wyjatki_z |
|---|---|---|---|
| `OPN` | `2,3` | | |
| `OPN_3Q` | `3,5` | `2,4` | `OPN` |

`wyjatki_z` sprawia, że wariant **dziedziczy wszystkie wyjątki spółki macierzystej**
(znaki kosztów, nietypowe klucze), więc nie powtarza się ich w arkuszu `Aliasy`.
Wariantów jest jeden na *układ*, nie na okres — Oponeo ma dwa i tyle zostanie.

Nazwa pliku wynikowego zawiera okres, żeby kolejny raport nie nadpisał poprzedniego:

| pierwszy nagłówek | plik |
|---|---|
| `I kw 2026 / 31.03.2026` | `wynik_OPN_1Q2026.xlsx` |
| `I polrocze 2026 / 30.06.2026` | `wynik_OPN_1H2026.xlsx` |
| `9 m-cy 2025 / 30.09.2025` | `wynik_WWL_9M2025.xlsx` |
| `I kw obrotowy 2025/26 / 30.06.2025` | `wynik_FTE_2025-06-30.xlsx` |

Przy **przesuniętym roku obrotowym** program nie zgaduje numeru kwartału, tylko wstawia
datę bilansową — `1Q2025` myliłoby się z kwartałem kalendarzowym.

## Jak dodać spółkę

1. `python start.py`, podaj nowy skrót → kreator wykryje układ pliku i zaproponuje aliasy
2. w `propozycja_TICKER.xlsx` popraw kolumnę `standard_key`
3. skopiuj wiersz z arkusza `1_Wiersz_do_Spolki` do arkusza `Spolki` w `slownik.xlsx`
4. skopiuj zatwierdzone aliasy (kolumny A–I) do arkusza `Aliasy`
5. wróć do okna programu, Enter — zmapuje i pokaże walidację

Powtarzaj aż `Niezmapowane` będzie puste, a walidacja pokaże komplet OK.

## Co program sprawdza

14 sum kontrolnych: aktywa = pasywa, składniki każdej sekcji bilansu, zysk ze sprzedaży,
EBIT, zysk przed opodatkowaniem, przepływy razem, kapitał obrotowy. Plus kontrola miękka —
implikowana stopa podatku poza 0–40% zwykle oznacza błąd mapowania wyniku.

Jeśli masz plik wzorcowy spółki w standardzie, program porówna z nim wynik pozycja po pozycji.
To najskuteczniejszy sposób wyłapywania błędów — przy każdej dotąd zmapowanej spółce
wywracał założenia, które wyglądały na oczywiste.

## Dwa szablony standardu

Arkusz `Spolki` ma kolumnę `szablon`:

- **`RAP`** — spółki przemysłowe i handlowe, 77 kluczy, 14 sum kontrolnych
- **`RAP_B`** — banki, 74 własne klucze (`interest_income`, `net_fee_income`, `loans_to_customers`...),
  13 sum kontrolnych. Bank nie ma kapitału obrotowego ani podziału na aktywa trwałe i obrotowe.

Klucze bankowe leżą w arkuszu `Klucze_B`.

## Źródła danych

| format | narzędzie | uwagi |
|---|---|---|
| scrapowany `.xlsx` | brak, wchodzi wprost | najpewniejszy — nazwy pełne |
| PDF | `pdf_na_xlsx.py` | nazwy łamane przez konwerter, ale są noty |
| XHTML (pdf2htmlEX) | `xhtml_na_xlsx.py` | to **nie** ESEF — brak tabel i znaczników XBRL |

Konfiguracja sekcji w konwerterach jest per spółka — przy nowym raporcie trzeba ją dostosować.

## Ustalona metodologia

- **kapitał obrotowy** nie jest przepisywany z rachunku przepływów, tylko liczony z różnic
  kolejnych bilansów (potwierdzone co do złotówki na trzech spółkach)
- `current_borrowings` / `noncurrent_borrowings` = **wyłącznie** linia „Kredyty i pożyczki";
  faktoring i „inne zobowiązania finansowe" idą w `*_other_liabilities`
- `property` zawiera prawo do użytkowania; `right_to_use_assets` to pozycja „w tym"
- w wariancie **porównawczym** całe „Koszty działalności operacyjnej" idą w `cost_of_sales`,
  a `distribution_expenses` i `administrative_expenses` zostają puste
- w ustawie o rachunkowości `reckoning` = rezerwy + rozliczenia międzyokresowe
- **kapitał obrotowy liczy się od końca POPRZEDNIEGO ROKU**, nie poprzedniego kwartału
  (przy I kwartale to to samo, przy II–IV już nie)
- rachunek wyników i przepływy są **narastające**; IV kwartał = cały rok. Gdy raport
  podaje obie wersje (III kwartał: osobno kwartał i osobno 9 miesięcy), bierz **narastająco**
- **raport roczny przekształca dane porównawcze**, standard trzyma bilans pierwotnie
  zaraportowany — przy raporcie rocznym dociągaj okres odniesienia z pliku wzorcowego
  (Oponeo 2025: raport pokazuje na 31.12.2024 należności 61 043, wzorzec 62 596)
- reguła kapitału obrotowego obowiązuje **od IV kwartału 2022** — starsze kolumny standardu
  liczone są inną metodą
- `year_profit` jest wypełniany tylko wtedy, gdy spółka wykazuje wynik okresu osobno w bilansie
- strata nie wymaga obsługi: klucze wynikowe wychodzą ujemne, koszty zostają dodatnie

Pełny opis w arkuszu `README` wewnątrz `slownik.xlsx`.

## Gdy coś nie działa

**`ModuleNotFoundError: No module named 'openpyxl'`** — brakuje biblioteki. Skrypty `.bat`
instalują ją same; jeśli uruchamiasz ręcznie, użyj `python -m pip install openpyxl`
(zapis `python -m pip` gwarantuje, że biblioteka trafi do tego Pythona, który faktycznie
uruchamia program — przy kilku wersjach Pythona samo `pip` potrafi trafić w inną).

**`python` nie jest rozpoznawane jako polecenie** — Python nie jest w PATH. Zainstaluj
z python.org zaznaczając „Add Python to PATH", albo używaj `py -3` zamiast `python`.

**Program pyta o plik, którego nie widać na liście** — wybierz `0` i wklej pełną ścieżkę.

## Uwaga o danych

Sprawozdania spółek i pliki wzorcowe trzymaj **poza repozytorium**, chyba że jest prywatne.
`.gitignore` domyślnie ignoruje tylko wyniki działania programu.

## Wersjonowanie słownika

`slownik.xlsx` jest binarny — `git diff` nic z niego nie pokaże, a konfliktu scalania nie da się
rozwiązać inaczej niż wyborem całej jednej wersji. Dlatego w repo **nie ma go wcale**:
źródłem prawdy są `slownik_csv/*.csv`, a xlsx jest z nich generowany.

```bash
git pull
python slownik_z_csv.py     # CSV  -> slownik.xlsx  (raz, po pobraniu zmian)
#  ... praca w Excelu, dopisywanie aliasów ...
python slownik_csv.py       # xlsx -> CSV           (przed commitem)
git add slownik_csv && git commit -m "AOL: aliasy ustawy o rachunkowości"
```

Dzięki temu w historii widać dokładnie, który alias się zmienił — i dwie osoby mogą
dopisywać aliasy równolegle bez tracenia pracy.

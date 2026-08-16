# Mapowanie sprawozdań na standard biznesradar

Zamienia plik sprawozdania spółki giełdowej na kolumnę w standardowym nazewnictwie
(`total_assets`, `revenues`, `ebit`, ...). Obsługuje MSSF i ustawę o rachunkowości.

## Szybki start

Na Windows kliknij **`start.bat`** — sam sprawdzi Pythona i doinstaluje `openpyxl`, jeśli brakuje.

Ręcznie:

```bash
python -m pip install -r requirements.txt
python start.py
```

Program pyta o skrót spółki, prosi o wskazanie plików z dysku i zapisuje `wynik_TICKER.xlsx`.

W Google Colab — w komórce notatnika (**nie** przez `!python start.py`, podproces nie ma klawiatury):

```python
import start
start.main()
```

## Pliki

| plik | rola |
|---|---|
| `slownik.xlsx` | **jedyny plik z wiedzą** — aliasy, reguły, konfiguracja spółek |
| `brmap.py` | cała logika (parser, mapowanie, walidacja, porównanie, kreator) |
| `start.py` | program interaktywny |
| `map.py`, `porownaj.py` | nakładki wiersza poleceń na `brmap.py` |
| `build_slownik.py` | generator słownika od zera — historia decyzji mapowania |
| `slownik_csv.py` | eksport słownika do CSV, żeby `git diff` był czytelny |
| `biznesradar_colab.ipynb` | notatnik Colab |

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

## Ustalona metodologia

- **kapitał obrotowy** nie jest przepisywany z rachunku przepływów, tylko liczony z różnic
  kolejnych bilansów (potwierdzone co do złotówki na trzech spółkach)
- `current_borrowings` / `noncurrent_borrowings` = **wyłącznie** linia „Kredyty i pożyczki";
  faktoring i „inne zobowiązania finansowe" idą w `*_other_liabilities`
- `property` zawiera prawo do użytkowania; `right_to_use_assets` to pozycja „w tym"
- w wariancie **porównawczym** całe „Koszty działalności operacyjnej" idą w `cost_of_sales`,
  a `distribution_expenses` i `administrative_expenses` zostają puste
- w ustawie o rachunkowości `reckoning` = rezerwy + rozliczenia międzyokresowe

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
rozwiązać inaczej niż wyborem jednej wersji. Przed commitem uruchom:

```bash
python slownik_csv.py
```

Zrzuca każdy arkusz do `slownik_csv/*.csv`. Wtedy w historii widać dokładnie, który alias się zmienił.

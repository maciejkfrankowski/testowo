#!/usr/bin/env bash
# Pierwsze wypchniecie na GitHub.
#   ./init_repo.sh https://github.com/UZYTKOWNIK/NAZWA.git
set -e
cd "$(dirname "$0")"

[ -z "$1" ] && { echo "Podaj adres pustego repozytorium, np.:"; \
  echo "   ./init_repo.sh https://github.com/macief/biznesradar-mapowanie.git"; exit 1; }

PY=$(command -v python3 || command -v python) || { echo "Nie znaleziono Pythona."; exit 1; }
echo "Python: $($PY --version)"
if ! $PY -c "import openpyxl" 2>/dev/null; then
  echo "Brakuje biblioteki openpyxl - instaluje..."
  $PY -m pip install --quiet -r requirements.txt || {
    echo "Instalacja nie powiodla sie. Sprobuj: $PY -m pip install openpyxl"; exit 1; }
fi
command -v git >/dev/null || { echo "Nie znaleziono gita."; exit 1; }

echo
echo "== kontrola przed wypchnieciem =="
$PY slownik_csv.py
$PY test_slownik.py

echo
git init -q 2>/dev/null || true
git add -A
git status --short
echo
read -p "Wypchnac powyzsze na $1 ? [t/N] " ODP
[[ "$ODP" =~ ^[TtYy] ]] || { echo "Przerwano."; exit 0; }

git commit -qm "Mapowanie sprawozdan na standard biznesradar"
git branch -M main
git remote remove origin 2>/dev/null || true
git remote add origin "$1"
git push -u origin main
echo
echo "Gotowe."

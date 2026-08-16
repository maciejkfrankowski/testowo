# -*- coding: utf-8 -*-
"""Nakladka wiersza polecen na brmap.py.

    python porownaj.py --wzorzec NEU.xlsx --data 2026-03-31 --wynik wynik_NEU.xlsx
Dla spolek raportujacych w zlotych (zaokraglenia +/-1) dodaj:  --tolerancja 2
"""
import argparse
import brmap

ap = argparse.ArgumentParser()
ap.add_argument("--wzorzec", required=True)
ap.add_argument("--data", required=True)
ap.add_argument("--wynik", required=True)
ap.add_argument("--out")
ap.add_argument("--arkusz", default="Spr Fin")
ap.add_argument("--kolumna", type=int, default=4)
ap.add_argument("--tolerancja", type=float, default=1.0)
args = ap.parse_args()

brmap.porownaj(args.wzorzec, args.data, args.wynik, args.out,
               args.arkusz, args.kolumna, tol=args.tolerancja)

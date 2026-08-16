# -*- coding: utf-8 -*-
"""Nakladka wiersza polecen na brmap.py (jedyne zrodlo logiki).

    python map.py --slownik slownik.xlsx --spolka NEU --plik "raport.xlsx" --out wynik_NEU.xlsx

Opcjonalnie, gdy kolumna porownawcza NIE jest poprzednim kwartalem (uklad UoR):
    --wzorzec AOL.xlsx --data-poprzednia 2025-12-31
"""
import argparse
import brmap

ap = argparse.ArgumentParser()
ap.add_argument("--slownik", required=True)
ap.add_argument("--spolka", required=True)
ap.add_argument("--plik", required=True)
ap.add_argument("--out")
ap.add_argument("--wzorzec", help="plik standardowy, z ktorego dociagnac bilans poprzedniego kwartalu")
ap.add_argument("--data-poprzednia", dest="data_poprzednia", help="RRRR-MM-DD")
args = ap.parse_args()

prev = None
if args.wzorzec and args.data_poprzednia:
    prev = brmap.wczytaj_okres_wzorca(args.wzorzec, args.data_poprzednia)

brmap.mapuj(args.slownik, args.spolka, args.plik, args.out, bilans_poprzedni=prev)

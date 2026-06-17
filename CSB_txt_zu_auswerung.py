#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ladeplan TXT auslesen
---------------------
Liest einen Tour- / Ladeplan als TXT-Datei ein und erzeugt eine saubere Liste
mit Tour, Wochentag, Ladefolge, CSB-Nummer, Name, Straße, Postleitzahl und Ort.

Aufruf:
    python ladeplan_txt_auslesen.py Ladeplan.txt

Ausgabe im gleichen Ordner:
    Ladeplan_Kunden.csv
    Ladeplan_Touren.csv
    Ladeplan_Pruefung.csv

Die CSV-Dateien werden mit Semikolon-Trennung und UTF-8-BOM geschrieben,
damit sie in Excel sauber geöffnet werden.
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Iterable, Optional


@dataclass
class KundeZeile:
    quelle: str
    druckdatum: str
    wochentag: str
    seite: str
    tour: str
    tour_text: str
    position_in_tour: int
    ladefolge: int
    ladefolge_aus_txt: str
    csb_nummer: str
    name: str
    strasse: str
    plz: str
    ort: str


@dataclass
class TourInfo:
    quelle: str
    druckdatum: str
    wochentag: str
    seite: str
    tour: str
    tour_text: str
    kunden_ausgelesen: int
    anzahl_kunden_im_txt: str
    pruefung: str


def read_text_smart(path: Path) -> tuple[str, str]:
    """Versucht typische Encodings. Der Ladeplan ist meist Windows-1252."""
    encodings = ["utf-8-sig", "utf-8", "cp1252", "latin1"]
    last_error: Optional[Exception] = None

    for encoding in encodings:
        try:
            return path.read_text(encoding=encoding), encoding
        except UnicodeDecodeError as exc:
            last_error = exc

    raise RuntimeError(f"Datei konnte nicht gelesen werden: {path} / {last_error}")


def clean(value: str) -> str:
    return " ".join(value.replace("\x00", " ").split())


def parse_ladeplan(text: str, source_name: str) -> tuple[list[KundeZeile], list[TourInfo]]:
    """
    Erkennt Tourblöcke und Kundenzeilen anhand der festen Spalten im TXT.

    Unterstützte Kundenzeilen:
      - ohne gedruckte Ladefolge vorne:
            10502 Name                 Straße                         21365 ORT
      - mit gedruckter Ladefolge vorne:
         1  13822 Name                 Straße                         24960 ORT

    Wichtige feste Spalten aus dem Ausdruck:
      - CSB steht ungefähr in Spalte 11 bis 16
      - PLZ steht ungefähr in Spalte 69 bis 73
    """
    rows: list[KundeZeile] = []
    tours: dict[tuple[str, str], dict[str, object]] = {}

    current_wochentag = ""
    current_tour = ""
    current_tour_text = ""
    current_seite = ""
    druckdatum = ""
    position = 0

    tour_header_re = re.compile(r"^\s*Tour\s+(\d{3,6})\s*(.*?)\s+LKW:")
    wochentag_re = re.compile(r"Wochentag\s+([A-Za-zÄÖÜäöüß]+)")
    seite_re = re.compile(r"\bSeite\s+(\d+)\b")
    druckdatum_re = re.compile(r"Druckdatum:\s*([0-9]{1,2}\.[0-9]{1,2}\.[0-9]{2,4})")
    anzahl_re = re.compile(r"^\s*(\d+)\s+Anzahl Kunden\b")

    def tour_key() -> tuple[str, str]:
        return current_wochentag, current_tour

    def ensure_tour() -> None:
        if not current_tour:
            return
        key = tour_key()
        if key not in tours:
            tours[key] = {
                "quelle": source_name,
                "druckdatum": druckdatum,
                "wochentag": current_wochentag,
                "seite": current_seite,
                "tour": current_tour,
                "tour_text": current_tour_text,
                "anzahl_kunden_im_txt": "",
                "kunden_ausgelesen": 0,
            }

    for raw_line in text.splitlines():
        line = raw_line.rstrip("\n\r")

        match = seite_re.search(line)
        if match:
            current_seite = match.group(1)

        match = druckdatum_re.search(line)
        if match:
            druckdatum = match.group(1)

        match = wochentag_re.search(line)
        if match:
            current_wochentag = match.group(1)

        match = tour_header_re.match(line)
        if match:
            current_tour = match.group(1).strip()
            current_tour_text = clean(match.group(2))
            position = 0
            ensure_tour()
            continue

        if current_tour:
            match = anzahl_re.match(line)
            if match:
                ensure_tour()
                tours[tour_key()]["anzahl_kunden_im_txt"] = match.group(1)
                continue

        # Kundenzeilen haben eine feste Breite. Die Erkennung ist bewusst streng,
        # damit Telefonnummern, Summen und Kopfzeilen nicht als Kunde erkannt werden.
        if not current_tour or len(line) < 74:
            continue

        ladefolge_feld = line[:10].strip()
        csb_feld = line[10:16].strip()
        plz_feld = line[68:73].strip()

        if not (csb_feld.isdigit() and plz_feld.isdigit() and len(plz_feld) == 5):
            continue

        position += 1
        ladefolge = int(ladefolge_feld) if ladefolge_feld.isdigit() else position

        ende_ort = line.find(" . . .", 73)
        if ende_ort == -1:
            ende_ort = min(len(line), 94)

        row = KundeZeile(
            quelle=source_name,
            druckdatum=druckdatum,
            wochentag=current_wochentag,
            seite=current_seite,
            tour=current_tour,
            tour_text=current_tour_text,
            position_in_tour=position,
            ladefolge=ladefolge,
            ladefolge_aus_txt=ladefolge_feld,
            csb_nummer=csb_feld,
            name=clean(line[16:37]),
            strasse=clean(line[37:68]),
            plz=plz_feld,
            ort=clean(line[74:ende_ort]),
        )
        rows.append(row)

        ensure_tour()
        tours[tour_key()]["kunden_ausgelesen"] = int(tours[tour_key()]["kunden_ausgelesen"]) + 1

    tour_infos: list[TourInfo] = []
    for key in sorted(tours.keys(), key=lambda item: (wochentag_sort(item[0]), int(item[1]) if item[1].isdigit() else item[1])):
        item = tours[key]
        ausgelesen = int(item["kunden_ausgelesen"])
        erwartet = str(item["anzahl_kunden_im_txt"])

        if erwartet == "":
            pruefung = "Keine Anzahl im TXT gefunden"
        elif erwartet.isdigit() and int(erwartet) == ausgelesen:
            pruefung = "OK"
        else:
            pruefung = "Abweichung"

        tour_infos.append(
            TourInfo(
                quelle=str(item["quelle"]),
                druckdatum=str(item["druckdatum"]),
                wochentag=str(item["wochentag"]),
                seite=str(item["seite"]),
                tour=str(item["tour"]),
                tour_text=str(item["tour_text"]),
                kunden_ausgelesen=ausgelesen,
                anzahl_kunden_im_txt=erwartet,
                pruefung=pruefung,
            )
        )

    return rows, tour_infos


def wochentag_sort(value: str) -> int:
    order = {
        "Montag": 1,
        "Dienstag": 2,
        "Mittwoch": 3,
        "Donnerstag": 4,
        "Freitag": 5,
        "Samstag": 6,
        "Sonntag": 7,
    }
    return order.get(value, 99)


def write_csv(path: Path, rows: Iterable[object]) -> int:
    rows = list(rows)
    if not rows:
        path.write_text("", encoding="utf-8-sig")
        return 0

    fieldnames = list(asdict(rows[0]).keys())
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames, delimiter=";")
        writer.writeheader()
        for row in rows:
            writer.writerow(asdict(row))
    return len(rows)



def main() -> int:
    parser = argparse.ArgumentParser(description="Liest Touren und CSB-Kunden aus einem Ladeplan TXT aus.")
    parser.add_argument("txt_datei", help="Pfad zur Ladeplan TXT-Datei")
    parser.add_argument(
        "--out",
        default="",
        help="Ausgabeordner. Standard: gleicher Ordner wie die TXT-Datei",
    )
    args = parser.parse_args()

    input_path = Path(args.txt_datei).expanduser().resolve()
    if not input_path.exists():
        print(f"Datei nicht gefunden: {input_path}", file=sys.stderr)
        return 2

    output_dir = Path(args.out).expanduser().resolve() if args.out else input_path.parent
    output_dir.mkdir(parents=True, exist_ok=True)

    text, encoding = read_text_smart(input_path)
    kunden, touren = parse_ladeplan(text, input_path.name)

    kunden_csv = output_dir / "Ladeplan_Kunden.csv"
    touren_csv = output_dir / "Ladeplan_Touren.csv"
    pruefung_csv = output_dir / "Ladeplan_Pruefung.csv"
    write_csv(kunden_csv, kunden)
    write_csv(touren_csv, touren)
    write_csv(pruefung_csv, [row for row in touren if row.pruefung != "OK"])

    anzahl_abweichungen = sum(1 for row in touren if row.pruefung != "OK")

    print("Fertig.")
    print(f"Eingelesene Datei: {input_path}")
    print(f"Erkanntes Encoding: {encoding}")
    print(f"Touren gefunden: {len(touren)}")
    print(f"Kundenzeilen gefunden: {len(kunden)}")
    print(f"Prüfhinweise: {anzahl_abweichungen}")
    print(f"CSV Kundenliste: {kunden_csv}")
    print(f"CSV Tourenübersicht: {touren_csv}")
    print(f"CSV Prüfung: {pruefung_csv}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

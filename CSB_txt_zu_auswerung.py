#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CSB TXT zu Auswertung
---------------------
Liest einen Tour- / Ladeplan als TXT-Datei ein und erzeugt eine saubere Liste
mit Tour, Wochentag, Ladefolge, CSB-Nummer, Name, Straße, Postleitzahl und Ort.

Diese Datei funktioniert auf zwei Arten:

1) Als Web-App mit Upload:
   python CSB_txt_zu_auswerung.py
   oder:
   uvicorn CSB_txt_zu_auswerung:app --host 0.0.0.0 --port 8501

2) Als Kommandozeilen-Skript:
   python CSB_txt_zu_auswerung.py Ladeplan.txt
   python CSB_txt_zu_auswerung.py Ladeplan.txt --out export

Ausgabe Kommandozeile:
   Ladeplan_Kunden.csv
   Ladeplan_Touren.csv
   Ladeplan_Pruefung.csv
   Ladeplan_Auswertung.xlsx
"""

from __future__ import annotations

import argparse
import csv
import html
import io
import re
import sys
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable, Optional

import pandas as pd
from fastapi import FastAPI, File, UploadFile
from fastapi.responses import HTMLResponse, PlainTextResponse, StreamingResponse


APP_TITLE = "CSB Ladeplan TXT Auswertung"


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


def clean(value: str) -> str:
    return " ".join(value.replace("\x00", " ").split())


def read_text_smart(path: Path) -> tuple[str, str]:
    data = path.read_bytes()
    return decode_text_smart(data)


def decode_text_smart(data: bytes) -> tuple[str, str]:
    """Versucht typische Encodings. Der Ladeplan ist meist Windows-1252."""
    encodings = ["utf-8-sig", "utf-8", "cp1252", "latin1"]
    last_error: Optional[Exception] = None

    for encoding in encodings:
        try:
            return data.decode(encoding), encoding
        except UnicodeDecodeError as exc:
            last_error = exc

    raise RuntimeError(f"Datei konnte nicht gelesen werden: {last_error}")


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


def parse_ladeplan(text: str, source_name: str) -> tuple[list[KundeZeile], list[TourInfo]]:
    """
    Erkennt Tourblöcke und Kundenzeilen anhand der festen Spalten im TXT.

    Erwarteter Ausdruck:
      Tour           1001 VAL                                 LKW:
      ...
             10502 V.BERGMANN LEBENSM.V V.STAUFFENBERGSTR.1A           21365 ADENDORF
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

        # Kundenzeilen haben feste Breite. Diese Erkennung ist bewusst streng,
        # damit Telefonzeilen, Summen und Kopfzeilen nicht als Kunden erkannt werden.
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

        rows.append(
            KundeZeile(
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
        )

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


def to_dataframe(rows: Iterable[object]) -> pd.DataFrame:
    rows = list(rows)
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame([asdict(row) for row in rows])


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


def build_excel_bytes(kunden: list[KundeZeile], touren: list[TourInfo], encoding: str) -> bytes:
    kunden_df = to_dataframe(kunden)
    touren_df = to_dataframe(touren)
    pruefung_df = touren_df[touren_df["pruefung"] != "OK"].copy() if not touren_df.empty else pd.DataFrame()

    summary = pd.DataFrame(
        [
            ["Encoding", encoding],
            ["Touren erkannt", len(touren)],
            ["Kundenzeilen erkannt", len(kunden)],
            ["Prüfhinweise", len(pruefung_df)],
        ],
        columns=["Kennzahl", "Wert"],
    )

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Zusammenfassung", index=False)
        kunden_df.to_excel(writer, sheet_name="Kunden", index=False)
        touren_df.to_excel(writer, sheet_name="Touren", index=False)
        pruefung_df.to_excel(writer, sheet_name="Pruefung", index=False)

        for sheet_name in writer.book.sheetnames:
            worksheet = writer.book[sheet_name]
            worksheet.freeze_panes = "A2"
            for column_cells in worksheet.columns:
                max_len = 0
                col_letter = column_cells[0].column_letter
                for cell in column_cells[:500]:
                    value = "" if cell.value is None else str(cell.value)
                    max_len = max(max_len, len(value))
                worksheet.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 45)

    output.seek(0)
    return output.getvalue()


def export_files(input_path: Path, output_dir: Path) -> dict[str, object]:
    text, encoding = read_text_smart(input_path)
    kunden, touren = parse_ladeplan(text, input_path.name)

    output_dir.mkdir(parents=True, exist_ok=True)
    kunden_csv = output_dir / "Ladeplan_Kunden.csv"
    touren_csv = output_dir / "Ladeplan_Touren.csv"
    pruefung_csv = output_dir / "Ladeplan_Pruefung.csv"
    excel_path = output_dir / "Ladeplan_Auswertung.xlsx"

    write_csv(kunden_csv, kunden)
    write_csv(touren_csv, touren)
    write_csv(pruefung_csv, [row for row in touren if row.pruefung != "OK"])
    excel_path.write_bytes(build_excel_bytes(kunden, touren, encoding))

    return {
        "encoding": encoding,
        "kunden": len(kunden),
        "touren": len(touren),
        "pruefhinweise": sum(1 for row in touren if row.pruefung != "OK"),
        "kunden_csv": kunden_csv,
        "touren_csv": touren_csv,
        "pruefung_csv": pruefung_csv,
        "excel": excel_path,
    }


app = FastAPI(title=APP_TITLE)


@app.get("/health", response_class=PlainTextResponse)
def health() -> str:
    return "OK"


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return """
<!doctype html>
<html lang="de">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>CSB Ladeplan TXT Auswertung</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 0; background: #111827; color: #f9fafb; }
    .wrap { max-width: 760px; margin: 48px auto; padding: 24px; }
    .card { background: #1f2937; border: 1px solid #374151; border-radius: 18px; padding: 28px; box-shadow: 0 16px 45px rgba(0,0,0,.25); }
    h1 { margin-top: 0; font-size: 28px; }
    p { color: #d1d5db; line-height: 1.5; }
    input[type=file] { width: 100%; padding: 14px; border: 1px dashed #6b7280; border-radius: 12px; background: #111827; color: #f9fafb; }
    button { margin-top: 18px; padding: 13px 18px; border: 0; border-radius: 12px; background: #f59e0b; color: #111827; font-weight: 700; cursor: pointer; }
    button:hover { filter: brightness(1.05); }
    .hint { margin-top: 18px; font-size: 14px; color: #9ca3af; }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card">
      <h1>CSB Ladeplan TXT Auswertung</h1>
      <p>Lade deinen Tour- / Ladeplan als TXT hoch. Danach bekommst du eine Excel mit Kunden, Touren und Prüfung.</p>
      <form action="/auswerten" method="post" enctype="multipart/form-data">
        <input type="file" name="file" accept=".txt,text/plain" required>
        <br>
        <button type="submit">Auswertung erstellen</button>
      </form>
      <div class="hint">Ergebnis: Ladeplan_Auswertung.xlsx</div>
    </div>
  </div>
</body>
</html>
"""


@app.post("/auswerten")
async def auswerten(file: UploadFile = File(...)) -> StreamingResponse:
    data = await file.read()
    text, encoding = decode_text_smart(data)
    filename = file.filename or "Ladeplan.txt"

    kunden, touren = parse_ladeplan(text, filename)
    excel_bytes = build_excel_bytes(kunden, touren, encoding)

    safe_name = re.sub(r"[^A-Za-z0-9_.-]+", "_", Path(filename).stem).strip("_") or "Ladeplan"
    download_name = f"{safe_name}_Auswertung.xlsx"

    headers = {
        "Content-Disposition": f"attachment; filename={html.escape(download_name)}"
    }
    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="Liest Touren und CSB-Kunden aus einem Ladeplan TXT aus.")
    parser.add_argument("txt_datei", nargs="?", help="Pfad zur Ladeplan TXT-Datei. Ohne Angabe startet die Web-App.")
    parser.add_argument("--out", default="", help="Ausgabeordner. Standard: gleicher Ordner wie die TXT-Datei")
    parser.add_argument("--host", default="0.0.0.0", help="Host für Web-App, Standard: 0.0.0.0")
    parser.add_argument("--port", type=int, default=8501, help="Port für Web-App, Standard: 8501")
    args = parser.parse_args()

    if not args.txt_datei:
        import uvicorn
        print(f"Starte Web-App: http://{args.host}:{args.port}")
        uvicorn.run("CSB_txt_zu_auswerung:app", host=args.host, port=args.port, reload=False)
        return 0

    input_path = Path(args.txt_datei).expanduser().resolve()
    if not input_path.exists():
        print(f"Datei nicht gefunden: {input_path}", file=sys.stderr)
        return 2

    output_dir = Path(args.out).expanduser().resolve() if args.out else input_path.parent
    result = export_files(input_path, output_dir)

    print("Fertig.")
    print(f"Eingelesene Datei: {input_path}")
    print(f"Erkanntes Encoding: {result['encoding']}")
    print(f"Touren gefunden: {result['touren']}")
    print(f"Kundenzeilen gefunden: {result['kunden']}")
    print(f"Prüfhinweise: {result['pruefhinweise']}")
    print(f"Excel: {result['excel']}")
    print(f"CSV Kundenliste: {result['kunden_csv']}")
    print(f"CSV Tourenübersicht: {result['touren_csv']}")
    print(f"CSV Prüfung: {result['pruefung_csv']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

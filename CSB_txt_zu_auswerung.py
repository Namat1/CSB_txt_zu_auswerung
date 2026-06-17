# CSB_txt_zu_auswerung.py
# Streamlit-App zum Auslesen eines CSB Tour-/Ladeplans aus einer TXT-Datei.
# Gibt Touren und Kunden mit CSB-Nummer als Excel und CSV aus.

from __future__ import annotations

import io
import re
import sys
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st


APP_TITLE = "CSB Ladeplan TXT auslesen"


def decode_bytes(data: bytes) -> str:
    """TXT-Datei robust dekodieren."""
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            pass
    return data.decode("utf-8", errors="replace")


def clean_text(value: str) -> str:
    """Mehrfach-Leerzeichen und Steuerzeichen bereinigen."""
    if value is None:
        return ""
    value = value.replace("\x0c", " ")
    value = value.replace(" ", "ü")  # häufiges Ersatzzeichen bei falsch kodierten Umlauten
    value = re.sub(r"\s+", " ", value)
    return value.strip(" \t\r\n.;")


def split_kunde_strasse(left_part: str) -> Tuple[str, str]:
    """
    Trennt den linken Teil der Kundenzeile in Kunde und Straße.

    Der Ausdruck ist ein alter Festbreiten-Ausdruck. Namen sind teilweise gekürzt,
    deshalb wird die Straße über typische Straßen-Schlüsselwörter erkannt.
    """
    text = clean_text(left_part)

    # Wenn im Ausdruck klare Spaltenabstände vorhanden sind, diese zuerst nutzen.
    parts = re.split(r"\s{2,}", text, maxsplit=1)
    if len(parts) == 2 and len(parts[0]) >= 3 and len(parts[1]) >= 3:
        return clean_text(parts[0]), clean_text(parts[1])

    # Sonst über typische Straßenmuster trennen.
    street_patterns = [
        r"\b[A-ZÄÖÜ][A-ZÄÖÜa-zäöüß\-.]*STR(?:ASSE|\.|)\b",
        r"\b[A-ZÄÖÜ][A-ZÄÖÜa-zäöüß\-.]*WEG\b",
        r"\b[A-ZÄÖÜ][A-ZÄÖÜa-zäöüß\-.]*PLATZ\b",
        r"\b[A-ZÄÖÜ][A-ZÄÖÜa-zäöüß\-.]*RING\b",
        r"\b[A-ZÄÖÜ][A-ZÄÖÜa-zäöüß\-.]*ALLEE\b",
        r"\b[A-ZÄÖÜ][A-ZÄÖÜa-zäöüß\-.]*CHAUSSEE\b",
        r"\b[A-ZÄÖÜ][A-ZÄÖÜa-zäöüß\-.]*DAMM\b",
        r"\b[A-ZÄÖÜ][A-ZÄÖÜa-zäöüß\-.]*DEICH\b",
        r"\b[A-ZÄÖÜ][A-ZÄÖÜa-zäöüß\-.]*MARKT\b",
        r"\b[A-ZÄÖÜ][A-ZÄÖÜa-zäöüß\-.]*LANDSTR(?:ASSE|\.|)\b",
        r"\bAM\b",
        r"\bAN DER\b",
        r"\bAUF DEM\b",
        r"\bBEI DER\b",
        r"\bIM\b",
        r"\bIN DER\b",
        r"\bZUM\b",
        r"\bZUR\b",
        r"\bV\.",
        r"\bDR\.",
        r"\bST\.-",
    ]

    best_pos = None
    for pat in street_patterns:
        match = re.search(pat, text)
        if match:
            pos = match.start()
            # Nicht direkt am Anfang trennen, sonst wird ein Kundenname fälschlich als Straße erkannt.
            if pos >= 8 and (best_pos is None or pos < best_pos):
                best_pos = pos

    if best_pos is not None:
        kunde = text[:best_pos].strip()
        strasse = text[best_pos:].strip()
        if kunde and strasse:
            return clean_text(kunde), clean_text(strasse)

    # Fallback: kein sauberer Trenner gefunden.
    return text, ""


def parse_ladeplan(text: str) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Liest Touren und Kunden aus dem CSB-TXT-Ladeplan.

    Ausgabe:
    - kunden_df: jede Kundenzeile
    - touren_df: je Tour eine Zusammenfassung
    - pruefung_df: Abgleich erkannte Kunden gegen "Anzahl Kunden" im Ausdruck
    """
    current_tour = ""
    current_tour_text = ""
    current_wochentag = ""
    position = 0

    rows: List[Dict[str, object]] = []
    tour_meta: Dict[str, Dict[str, object]] = {}

    lines = text.splitlines()

    tour_re = re.compile(r"^\s*Tour\s+(\d{3,6})\b(.*?)(?:LKW:|$)", re.IGNORECASE)
    day_re = re.compile(r"^\s*Wochentag\s+(.+?)(?:Fahrer:|$)", re.IGNORECASE)

    # Beispiel Kundenzeile:
    #          10502 V.BERGMANN ... V.STAUFFENBERGSTR.1A           21365 ADENDORF            . . . . . . .
    cust_re = re.compile(
        r"^\s{3,}(\d{3,6})\s+(.+?)\s+(\d{5})\s+(.+?)\s+(?:\.\s*){2,}\s*$"
    )

    count_re = re.compile(r"^\s*(\d+)\s+Anzahl Kunden\b", re.IGNORECASE)

    for raw_line in lines:
        line = raw_line.rstrip("\n")

        day_match = day_re.search(line)
        if day_match:
            current_wochentag = clean_text(day_match.group(1))

        tour_match = tour_re.search(line)
        if tour_match:
            current_tour = tour_match.group(1)
            current_tour_text = clean_text(tour_match.group(2))
            position = 0

            tour_meta.setdefault(
                current_tour,
                {
                    "Tour": current_tour,
                    "Wochentag": current_wochentag,
                    "Tour_Text": current_tour_text,
                    "Erwartete_Kunden": None,
                },
            )
            if current_wochentag:
                tour_meta[current_tour]["Wochentag"] = current_wochentag
            if current_tour_text:
                tour_meta[current_tour]["Tour_Text"] = current_tour_text
            continue

        count_match = count_re.search(line)
        if count_match and current_tour:
            tour_meta.setdefault(
                current_tour,
                {
                    "Tour": current_tour,
                    "Wochentag": current_wochentag,
                    "Tour_Text": current_tour_text,
                    "Erwartete_Kunden": None,
                },
            )
            tour_meta[current_tour]["Erwartete_Kunden"] = int(count_match.group(1))
            continue

        cust_match = cust_re.search(line)
        if cust_match and current_tour:
            position += 1

            csb = cust_match.group(1)
            left_part = cust_match.group(2)
            plz = cust_match.group(3)
            ort = clean_text(cust_match.group(4))

            kunde, strasse = split_kunde_strasse(left_part)

            rows.append(
                {
                    "Tour": current_tour,
                    "Wochentag": current_wochentag,
                    "Position_im_Tourblock": position,
                    "CSB": csb,
                    "Kunde": kunde,
                    "Strasse": strasse,
                    "PLZ": plz,
                    "Ort": ort,
                    "Tour_Text": current_tour_text,
                    "Originalzeile": clean_text(line),
                }
            )

    kunden_df = pd.DataFrame(rows)

    if kunden_df.empty:
        touren_df = pd.DataFrame(columns=["Tour", "Wochentag", "Tour_Text", "Erwartete_Kunden", "Erkannte_Kunden", "Differenz"])
        pruefung_df = pd.DataFrame(columns=["Tour", "Wochentag", "Erwartete_Kunden", "Erkannte_Kunden", "Differenz", "Status"])
        return kunden_df, touren_df, pruefung_df

    erkannte = (
        kunden_df.groupby("Tour", as_index=False)
        .size()
        .rename(columns={"size": "Erkannte_Kunden"})
    )

    touren_df = pd.DataFrame(tour_meta.values())
    touren_df = touren_df.merge(erkannte, on="Tour", how="outer")
    touren_df["Erwartete_Kunden"] = pd.to_numeric(touren_df["Erwartete_Kunden"], errors="coerce")
    touren_df["Erkannte_Kunden"] = pd.to_numeric(touren_df["Erkannte_Kunden"], errors="coerce").fillna(0).astype(int)
    touren_df["Differenz"] = touren_df["Erkannte_Kunden"] - touren_df["Erwartete_Kunden"]

    def status(row) -> str:
        if pd.isna(row["Erwartete_Kunden"]):
            return "Keine Sollzahl gefunden"
        if row["Differenz"] == 0:
            return "OK"
        return "Abweichung"

    touren_df["Status"] = touren_df.apply(status, axis=1)
    touren_df = touren_df.sort_values("Tour", kind="stable").reset_index(drop=True)

    pruefung_df = touren_df[["Tour", "Wochentag", "Erwartete_Kunden", "Erkannte_Kunden", "Differenz", "Status"]].copy()

    kunden_df = kunden_df.sort_values(["Tour", "Position_im_Tourblock"], kind="stable").reset_index(drop=True)
    return kunden_df, touren_df, pruefung_df


def make_excel(kunden_df: pd.DataFrame, touren_df: pd.DataFrame, pruefung_df: pd.DataFrame) -> bytes:
    """Erstellt eine Excel-Datei im Speicher."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        kunden_df.to_excel(writer, sheet_name="Kunden", index=False)
        touren_df.to_excel(writer, sheet_name="Touren", index=False)
        pruefung_df.to_excel(writer, sheet_name="Pruefung", index=False)

        # einfache Spaltenbreiten
        for sheet_name in writer.book.sheetnames:
            ws = writer.book[sheet_name]
            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    value = "" if cell.value is None else str(cell.value)
                    max_len = max(max_len, len(value))
                ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 55)

    return output.getvalue()


def run_streamlit_app() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="🚚", layout="wide")

    st.title("🚚 CSB Ladeplan TXT auslesen")
    st.caption("TXT hochladen → Touren und Kunden mit CSB-Nummer als Excel oder CSV exportieren.")

    uploaded = st.file_uploader("Ladeplan als TXT hochladen", type=["txt"])

    if uploaded is None:
        st.info("Bitte eine TXT-Datei hochladen.")
        st.stop()

    text = decode_bytes(uploaded.getvalue())
    kunden_df, touren_df, pruefung_df = parse_ladeplan(text)

    if kunden_df.empty:
        st.error("Es wurden keine Kundenzeilen erkannt. Bitte prüfen, ob es wirklich ein CSB Tour-/Ladeplan als TXT ist.")
        with st.expander("Vorschau der Datei"):
            st.text(text[:5000])
        st.stop()

    fehler = pruefung_df[pruefung_df["Status"] != "OK"]

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Touren", f"{touren_df['Tour'].nunique():,}".replace(",", "."))
    col2.metric("Kundenzeilen", f"{len(kunden_df):,}".replace(",", "."))
    col3.metric("Prüfabweichungen", f"{len(fehler):,}".replace(",", "."))
    col4.metric("Datei", uploaded.name)

    if len(fehler) == 0:
        st.success("Alle Touren passen zur im Ausdruck angegebenen Anzahl Kunden.")
    else:
        st.warning("Es gibt Touren, bei denen die erkannte Anzahl nicht zur Sollzahl im Ausdruck passt.")

    excel_bytes = make_excel(kunden_df, touren_df, pruefung_df)
    csv_bytes = kunden_df.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")

    d1, d2 = st.columns(2)
    d1.download_button(
        "Excel herunterladen",
        data=excel_bytes,
        file_name="Ladeplan_Auswertung.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    d2.download_button(
        "Kunden CSV herunterladen",
        data=csv_bytes,
        file_name="Ladeplan_Kunden.csv",
        mime="text/csv",
        use_container_width=True,
    )

    tab1, tab2, tab3 = st.tabs(["Kunden", "Touren", "Prüfung"])

    with tab1:
        st.dataframe(kunden_df, use_container_width=True, hide_index=True)

    with tab2:
        st.dataframe(touren_df, use_container_width=True, hide_index=True)

    with tab3:
        st.dataframe(pruefung_df, use_container_width=True, hide_index=True)


def run_cli(txt_path: Path, out_dir: Path) -> None:
    text = decode_bytes(txt_path.read_bytes())
    kunden_df, touren_df, pruefung_df = parse_ladeplan(text)

    out_dir.mkdir(parents=True, exist_ok=True)
    kunden_df.to_csv(out_dir / "Ladeplan_Kunden.csv", index=False, sep=";", encoding="utf-8-sig")
    touren_df.to_csv(out_dir / "Ladeplan_Touren.csv", index=False, sep=";", encoding="utf-8-sig")
    pruefung_df.to_csv(out_dir / "Ladeplan_Pruefung.csv", index=False, sep=";", encoding="utf-8-sig")
    (out_dir / "Ladeplan_Auswertung.xlsx").write_bytes(make_excel(kunden_df, touren_df, pruefung_df))

    print(f"Fertig: {len(touren_df)} Touren, {len(kunden_df)} Kunden")
    print(f"Ausgabeordner: {out_dir.resolve()}")


# Wichtig:
# Kein argparse mit Pflichtargument, weil Streamlit Cloud die Datei ohne TXT-Argument startet.
if len(sys.argv) >= 2 and Path(sys.argv[1]).is_file():
    output_dir = Path(sys.argv[2]) if len(sys.argv) >= 3 else Path("ausgabe")
    run_cli(Path(sys.argv[1]), output_dir)
else:
    run_streamlit_app()

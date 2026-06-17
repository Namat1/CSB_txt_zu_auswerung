# CSB_txt_zu_auswerung.py
# Streamlit-App: CSB Tour-/Ladeplan TXT sauber auslesen
# Ausgabe: Touren + Kunden mit CSB-Nummer, Name, Straße, Postleitzahl und Ort

from __future__ import annotations

import io
import re
import sys
from pathlib import Path

import pandas as pd
import streamlit as st


APP_TITLE = "CSB Ladeplan TXT auslesen"


def decode_bytes(data: bytes) -> str:
    """TXT-Datei robust dekodieren. CSB-Exporte enthalten oft Windows-1252-Umlaute."""
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("cp1252", errors="replace")


def clean_text(value: str) -> str:
    """Text bereinigen, ohne echte Umlaute zu zerstören."""
    if value is None:
        return ""
    value = value.replace("\x0c", " ")
    value = value.replace("\xa0", " ")
    value = value.replace("\x81", " ")
    value = re.sub(r"\s+", " ", value)
    return value.strip(" \t\r\n.;")


def extract_customer_line(line: str):
    """
    Liest eine Kundenzeile aus dem alten Festbreiten-Ausdruck.

    Unterstützt beide Varianten:
    - ohne Ladepositionsdruck:        10502 Kunde Straße PLZ Ort ...
    - mit Ladepositionsdruck:    1    13822 Kunde Straße PLZ Ort ...
    """
    raw = line.rstrip("\r\n").replace("\xa0", " ")

    # Kundenzeilen haben am Ende die Punkt-Spalten.
    if not re.search(r"(?:\.\s*){2,}\s*$", raw):
        return None

    # Variante mit optionaler Ladepositionsnummer vor der CSB-Nummer.
    m_start = re.match(r"^\s{3,}(?:(\d{1,3})\s+)?(\d{3,6})\s+", raw)
    if not m_start:
        return None

    ladefolge_gedruckt = m_start.group(1)
    csb = m_start.group(2)

    # Letzte fünfstellige Nummer vor dem Ort ist die Postleitzahl.
    plz_matches = list(re.finditer(r"\b\d{5}\b", raw))
    if not plz_matches:
        return None

    plz_match = plz_matches[-1]
    plz = plz_match.group(0)

    # Ort steht nach der Postleitzahl bis vor Punkt-Spalten.
    right = raw[plz_match.end():]
    right = re.sub(r"(?:\s+\.){2,}.*$", "", right)
    ort = clean_text(right)

    # Bereich zwischen CSB und PLZ enthält Name + Straße.
    mid = raw[m_start.end():plz_match.start()].rstrip()

    # Im Ausdruck ist der Kundenname 21 Zeichen breit, danach beginnt Straße.
    kunde = clean_text(mid[:21])
    strasse = clean_text(mid[21:])

    return ladefolge_gedruckt, csb, kunde, strasse, plz, ort


def parse_ladeplan(text: str):
    current_tour = ""
    current_tour_text = ""
    current_wochentag = ""
    position = 0

    rows = []
    tour_meta = {}

    tour_re = re.compile(r"^\s*Tour\s+(\d{3,6})\b(.*?)(?:LKW:|$)", re.IGNORECASE)
    day_re = re.compile(r"^\s*Wochentag\s+(.+?)(?:Fahrer:|$)", re.IGNORECASE)
    count_re = re.compile(r"^\s*(\d+)\s+Anzahl Kunden\b", re.IGNORECASE)

    for raw_line in text.splitlines():
        line = raw_line.rstrip("\n\r")

        day_match = day_re.search(line)
        if day_match:
            current_wochentag = clean_text(day_match.group(1))

        tour_match = tour_re.search(line)
        if tour_match:
            current_tour = tour_match.group(1)
            current_tour_text = clean_text(tour_match.group(2))
            position = 0

            tour_meta[current_tour] = {
                "Tour": current_tour,
                "Wochentag": current_wochentag,
                "Tour_Text": current_tour_text,
                "Erwartete_Kunden": None,
            }
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

        customer = extract_customer_line(line)
        if customer and current_tour:
            position += 1
            ladefolge_gedruckt, csb, kunde, strasse, plz, ort = customer

            ladereihenfolge = int(ladefolge_gedruckt) if ladefolge_gedruckt else position

            rows.append(
                {
                    "Tour": current_tour,
                    "Wochentag": current_wochentag,
                    "Ladereihenfolge": ladereihenfolge,
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
        empty_touren = pd.DataFrame(
            columns=["Tour", "Wochentag", "Tour_Text", "Erwartete_Kunden", "Erkannte_Kunden", "Differenz", "Status"]
        )
        empty_pruefung = pd.DataFrame(
            columns=["Tour", "Wochentag", "Erwartete_Kunden", "Erkannte_Kunden", "Differenz", "Status"]
        )
        return kunden_df, empty_touren, empty_pruefung

    erkannte = (
        kunden_df.groupby("Tour", as_index=False)
        .size()
        .rename(columns={"size": "Erkannte_Kunden"})
    )

    touren_df = pd.DataFrame(tour_meta.values()).merge(erkannte, on="Tour", how="outer")
    touren_df["Erwartete_Kunden"] = pd.to_numeric(touren_df["Erwartete_Kunden"], errors="coerce")
    touren_df["Erkannte_Kunden"] = pd.to_numeric(touren_df["Erkannte_Kunden"], errors="coerce").fillna(0).astype(int)
    touren_df["Differenz"] = touren_df["Erkannte_Kunden"] - touren_df["Erwartete_Kunden"]

    def make_status(row):
        if pd.isna(row["Erwartete_Kunden"]):
            return "Keine Sollzahl gefunden"
        if row["Differenz"] == 0:
            return "OK"
        return "Abweichung"

    touren_df["Status"] = touren_df.apply(make_status, axis=1)

    kunden_df = kunden_df.sort_values(["Tour", "Position_im_Tourblock"], kind="stable").reset_index(drop=True)
    touren_df = touren_df.sort_values("Tour", kind="stable").reset_index(drop=True)
    pruefung_df = touren_df[["Tour", "Wochentag", "Erwartete_Kunden", "Erkannte_Kunden", "Differenz", "Status"]].copy()

    return kunden_df, touren_df, pruefung_df


def make_excel(kunden_df, touren_df, pruefung_df) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        kunden_df.to_excel(writer, sheet_name="Kunden", index=False)
        touren_df.to_excel(writer, sheet_name="Touren", index=False)
        pruefung_df.to_excel(writer, sheet_name="Pruefung", index=False)

        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                max_len = 0
                letter = col[0].column_letter
                for cell in col:
                    value = "" if cell.value is None else str(cell.value)
                    max_len = max(max_len, len(value))
                ws.column_dimensions[letter].width = min(max(max_len + 2, 10), 55)

    return output.getvalue()


def run_streamlit_app() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="🚚", layout="wide")

    st.title("🚚 CSB Ladeplan TXT auslesen")
    st.caption("TXT hochladen → Touren und Kunden mit CSB-Nummer sauber als Excel oder CSV exportieren.")

    uploaded = st.file_uploader("Ladeplan als TXT hochladen", type=["txt"])

    if uploaded is None:
        st.info("Bitte eine TXT-Datei hochladen.")
        st.stop()

    text = decode_bytes(uploaded.getvalue())
    kunden_df, touren_df, pruefung_df = parse_ladeplan(text)

    if kunden_df.empty:
        st.error("Es wurden keine Kundenzeilen erkannt.")
        with st.expander("Dateivorschau"):
            st.text(text[:5000])
        st.stop()

    fehler = pruefung_df[pruefung_df["Status"] != "OK"]

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Touren", f"{touren_df['Tour'].nunique():,}".replace(",", "."))
    col2.metric("Kundenzeilen", f"{len(kunden_df):,}".replace(",", "."))
    col3.metric("Prüfabweichungen", f"{len(fehler):,}".replace(",", "."))
    col4.metric("Datei", uploaded.name)

    if len(fehler) == 0:
        st.success("Alle Touren passen zur angegebenen Anzahl Kunden.")
    else:
        st.warning("Es gibt Touren mit abweichender Kundenanzahl. Bitte im Reiter Prüfung ansehen.")

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

    print(f"Fertig: {touren_df['Tour'].nunique()} Touren, {len(kunden_df)} Kunden")
    print(f"Ausgabeordner: {out_dir.resolve()}")


# Streamlit Cloud startet ohne Startargument.
# Kommandozeile funktioniert trotzdem: python CSB_txt_zu_auswerung.py Ladeplan.txt
if len(sys.argv) >= 2 and Path(sys.argv[1]).is_file():
    output_dir = Path(sys.argv[2]) if len(sys.argv) >= 3 else Path("ausgabe")
    run_cli(Path(sys.argv[1]), output_dir)
else:
    run_streamlit_app()

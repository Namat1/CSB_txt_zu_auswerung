# CSB_txt_zu_auswerung.py
# Reine Streamlit-App ohne argparse, ohne FastAPI, ohne Pflichtargumente.
# Wichtig: Es wird KEINE Ladereihenfolge erfunden.
# Wenn im TXT vorne in der Spalte La. ein Wert steht, landet er in La_aus_TXT.
# Wenn dort nichts steht, bleibt La_aus_TXT leer.

from __future__ import annotations

import io
import re
import traceback

import streamlit as st


st.set_page_config(
    page_title="CSB Ladeplan TXT auslesen",
    page_icon="🚚",
    layout="wide",
)

st.title("🚚 CSB Ladeplan TXT auslesen")
st.caption("TXT hochladen → Touren und Kunden mit CSB-Nummer sauber als Excel oder CSV exportieren.")
st.info("Hinweis: Es wird keine Ladereihenfolge erfunden. Leere La.-Spalten bleiben leer.")


def decode_bytes(data: bytes) -> str:
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("cp1252", errors="replace")


def clean_text(value: str) -> str:
    if value is None:
        return ""
    value = value.replace("\x0c", " ")
    value = value.replace("\xa0", " ")
    value = value.replace("\x81", " ")
    value = re.sub(r"\s+", " ", value)
    return value.strip(" \t\r\n.;")


def extract_customer_line(line: str):
    """
    Liest eine Kundenzeile aus dem CSB-Ladeplan.

    Zwei Varianten werden erkannt:
    1. Ohne La.-Wert:
       10502 Kunde Straße PLZ Ort ...
       -> La_aus_TXT bleibt leer.

    2. Mit La.-Wert vorne:
       1    13822 Kunde Straße PLZ Ort ...
       -> La_aus_TXT = 1.

    Wichtig: Wenn kein La.-Wert vorhanden ist, wird NICHT automatisch 1,2,3 gezählt.
    """
    raw = line.rstrip("\r\n").replace("\xa0", " ")

    # Kundenzeilen enden im Ausdruck mit mehreren Punkt-Spalten.
    if not re.search(r"(?:\.\s*){2,}\s*$", raw):
        return None

    # Erst versuchen: La.-Wert plus CSB.
    # Beispiel: "     1    13822 Peter Saur..."
    match_with_la = re.match(r"^\s{3,}(\d{1,3})\s{2,}(\d{3,6})\s+", raw)

    if match_with_la:
        la_aus_txt = match_with_la.group(1)
        csb = match_with_la.group(2)
        start_end = match_with_la.end()
    else:
        # Ohne La.-Wert: erste Nummer ist CSB.
        # Beispiel: "          10502 V.BERGMANN..."
        match_without_la = re.match(r"^\s{3,}(\d{3,6})\s+", raw)
        if not match_without_la:
            return None
        la_aus_txt = ""
        csb = match_without_la.group(1)
        start_end = match_without_la.end()

    # Letzte fünfstellige Nummer vor dem Ort ist die Postleitzahl.
    plz_matches = list(re.finditer(r"\b\d{5}\b", raw))
    if not plz_matches:
        return None

    plz_match = plz_matches[-1]
    plz = plz_match.group(0)

    ort_raw = raw[plz_match.end():]
    ort_raw = re.sub(r"(?:\s+\.){2,}.*$", "", ort_raw)
    ort = clean_text(ort_raw)

    # Bereich zwischen CSB und PLZ enthält Name + Straße.
    mid = raw[start_end:plz_match.start()].rstrip()

    # CSB-Festbreite: Kundenname etwa 21 Zeichen, danach Straße.
    kunde = clean_text(mid[:21])
    strasse = clean_text(mid[21:])

    return la_aus_txt, csb, kunde, strasse, plz, ort


def parse_ladeplan(text: str):
    import pandas as pd

    current_tour = ""
    current_wochentag = ""
    current_tour_text = ""
    position = 0

    kunden_rows = []
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
            la_aus_txt, csb, kunde, strasse, plz, ort = customer

            kunden_rows.append(
                {
                    "Tour": current_tour,
                    "Wochentag": current_wochentag,
                    "CSB": csb,
                    "Kunde": kunde,
                    "Strasse": strasse,
                    "PLZ": plz,
                    "Ort": ort,
                    "La_aus_TXT": la_aus_txt,
                    "Position_im_Tourblock": position,
                    "Tour_Text": current_tour_text,
                }
            )

    kunden_df = pd.DataFrame(kunden_rows)

    if kunden_df.empty:
        touren_df = pd.DataFrame(
            columns=["Tour", "Wochentag", "Tour_Text", "Erwartete_Kunden", "Erkannte_Kunden", "Differenz", "Status"]
        )
        pruefung_df = pd.DataFrame(
            columns=["Tour", "Wochentag", "Erwartete_Kunden", "Erkannte_Kunden", "Differenz", "Status"]
        )
        return kunden_df, touren_df, pruefung_df

    erkannte = (
        kunden_df.groupby("Tour", as_index=False)
        .size()
        .rename(columns={"size": "Erkannte_Kunden"})
    )

    touren_df = pd.DataFrame(tour_meta.values()).merge(erkannte, on="Tour", how="outer")
    touren_df["Erwartete_Kunden"] = pd.to_numeric(touren_df["Erwartete_Kunden"], errors="coerce")
    touren_df["Erkannte_Kunden"] = pd.to_numeric(touren_df["Erkannte_Kunden"], errors="coerce").fillna(0).astype(int)
    touren_df["Differenz"] = touren_df["Erkannte_Kunden"] - touren_df["Erwartete_Kunden"]

    def status(row):
        if pd.isna(row["Erwartete_Kunden"]):
            return "Keine Sollzahl gefunden"
        if row["Differenz"] == 0:
            return "OK"
        return "Abweichung"

    touren_df["Status"] = touren_df.apply(status, axis=1)

    kunden_df = kunden_df.sort_values(["Tour", "Position_im_Tourblock"], kind="stable").reset_index(drop=True)
    touren_df = touren_df.sort_values("Tour", kind="stable").reset_index(drop=True)
    pruefung_df = touren_df[["Tour", "Wochentag", "Erwartete_Kunden", "Erkannte_Kunden", "Differenz", "Status"]].copy()

    return kunden_df, touren_df, pruefung_df


def make_excel(kunden_df, touren_df, pruefung_df) -> bytes:
    import pandas as pd

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        kunden_df.to_excel(writer, sheet_name="Kunden", index=False)
        touren_df.to_excel(writer, sheet_name="Touren", index=False)
        pruefung_df.to_excel(writer, sheet_name="Pruefung", index=False)

        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                letter = col[0].column_letter
                max_len = max(len("" if cell.value is None else str(cell.value)) for cell in col)
                ws.column_dimensions[letter].width = min(max(max_len + 2, 10), 55)

    return output.getvalue()


uploaded = st.file_uploader("Ladeplan als TXT hochladen", type=["txt"])

if uploaded is None:
    st.stop()

try:
    with st.spinner("Ladeplan wird gelesen..."):
        text = decode_bytes(uploaded.getvalue())
        kunden_df, touren_df, pruefung_df = parse_ladeplan(text)

    if kunden_df.empty:
        st.error("Es wurden keine Kundenzeilen erkannt.")
        with st.expander("Dateivorschau"):
            st.text(text[:5000])
        st.stop()

    fehler_df = pruefung_df[pruefung_df["Status"] != "OK"]
    la_gefunden = kunden_df["La_aus_TXT"].astype(str).str.strip().ne("").sum()

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Touren", f"{touren_df['Tour'].nunique():,}".replace(",", "."))
    col2.metric("Kunden", f"{len(kunden_df):,}".replace(",", "."))
    col3.metric("La.-Werte aus TXT", f"{la_gefunden:,}".replace(",", "."))
    col4.metric("Prüfabweichungen", f"{len(fehler_df):,}".replace(",", "."))

    if len(fehler_df) == 0:
        st.success("Alle Touren passen zur angegebenen Anzahl Kunden.")
    else:
        st.warning("Es gibt Touren mit abweichender Kundenanzahl. Bitte Reiter Prüfung ansehen.")

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

except Exception:
    st.error("Fehler beim Verarbeiten der Datei.")
    st.code(traceback.format_exc())

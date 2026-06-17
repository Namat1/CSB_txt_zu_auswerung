# CSB_txt_zu_auswerung.py
# Reine Streamlit-App ohne argparse, ohne FastAPI, ohne Pflichtargumente.

from __future__ import annotations

import io
import re

import pandas as pd
import streamlit as st


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
    raw = line.rstrip("\r\n").replace("\xa0", " ")

    # Kundenzeilen enden im Ausdruck mit mehreren Punkt-Spalten.
    if not re.search(r"(?:\.\s*){2,}\s*$", raw):
        return None

    # Unterstützt:
    # 10502 Kunde ...
    # 1 13822 Kunde ...
    match_start = re.match(r"^\s{3,}(?:(\d{1,3})\s+)?(\d{3,6})\s+", raw)
    if not match_start:
        return None

    ladefolge_gedruckt = match_start.group(1)
    csb = match_start.group(2)

    plz_matches = list(re.finditer(r"\b\d{5}\b", raw))
    if not plz_matches:
        return None

    plz_match = plz_matches[-1]
    plz = plz_match.group(0)

    ort_raw = raw[plz_match.end():]
    ort_raw = re.sub(r"(?:\s+\.){2,}.*$", "", ort_raw)
    ort = clean_text(ort_raw)

    mid = raw[match_start.end():plz_match.start()].rstrip()

    # CSB-Festbreite: Name 21 Zeichen, danach Straße.
    kunde = clean_text(mid[:21])
    strasse = clean_text(mid[21:])

    return ladefolge_gedruckt, csb, kunde, strasse, plz, ort


def parse_ladeplan(text: str):
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
            ladefolge_gedruckt, csb, kunde, strasse, plz, ort = customer

            kunden_rows.append(
                {
                    "Tour": current_tour,
                    "Wochentag": current_wochentag,
                    "Ladereihenfolge": int(ladefolge_gedruckt) if ladefolge_gedruckt else position,
                    "Position_im_Tourblock": position,
                    "CSB": csb,
                    "Kunde": kunde,
                    "Strasse": strasse,
                    "PLZ": plz,
                    "Ort": ort,
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


def make_excel(kunden_df: pd.DataFrame, touren_df: pd.DataFrame, pruefung_df: pd.DataFrame) -> bytes:
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


st.set_page_config(page_title="CSB Ladeplan TXT auslesen", page_icon="🚚", layout="wide")

st.title("🚚 CSB Ladeplan TXT auslesen")
st.caption("TXT hochladen → Touren und Kunden mit CSB-Nummer sauber als Excel oder CSV exportieren.")

uploaded = st.file_uploader("Ladeplan als TXT hochladen", type=["txt"])

if uploaded is None:
    st.info("Bitte eine TXT-Datei hochladen.")
    st.stop()

try:
    text = decode_bytes(uploaded.getvalue())
    kunden_df, touren_df, pruefung_df = parse_ladeplan(text)

    if kunden_df.empty:
        st.error("Es wurden keine Kundenzeilen erkannt.")
        with st.expander("Dateivorschau"):
            st.text(text[:5000])
        st.stop()

    fehler_df = pruefung_df[pruefung_df["Status"] != "OK"]

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Touren", f"{touren_df['Tour'].nunique():,}".replace(",", "."))
    col2.metric("Kunden", f"{len(kunden_df):,}".replace(",", "."))
    col3.metric("Prüfabweichungen", f"{len(fehler_df):,}".replace(",", "."))
    col4.metric("Datei", uploaded.name)

    if len(fehler_df) == 0:
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

except Exception as error:
    st.exception(error)

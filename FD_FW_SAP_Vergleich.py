import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="FD/FW Liefertag Vergleich", layout="wide")
st.title("FD / FW Liefertag Vergleich")

WOCHENTAGE = {1: "Mo", 2: "Di", 3: "Mi", 4: "Do", 5: "Fr", 6: "Sa", 7: "So"}

def load_file(file):
    df = pd.read_excel(file, header=0, dtype={0: str})
    df.columns = [str(c) for c in df.columns]
    # Spalten umbenennen auf Positionen
    cols = df.columns.tolist()
    df = df.rename(columns={
        cols[0]: "SAP",
        cols[1]: "Name",
        cols[6]: "Liefertag"
    })
    df["SAP"] = df["SAP"].astype(str).str.strip().str.zfill(0)
    df["Liefertag"] = pd.to_numeric(df["Liefertag"], errors="coerce")
    return df[["SAP", "Name", "Liefertag"]].dropna(subset=["Liefertag"])

def tage_str(tage_set):
    sorted_tage = sorted([int(t) for t in tage_set if pd.notna(t)])
    return ", ".join([f"{t}={WOCHENTAGE.get(t, '?')}" for t in sorted_tage])

col1, col2 = st.columns(2)
with col1:
    fd_file = st.file_uploader("FD-Datei (.xlsx)", type=["xlsx"], key="fd")
with col2:
    fw_file = st.file_uploader("FW-Datei (.xlsx)", type=["xlsx"], key="fw")

st.divider()

sap_input = st.text_area(
    "SAP-Nummern filtern (eine pro Zeile, leer = alle)",
    height=120,
    placeholder="12823\n12920\n12923"
)

run = st.button("Vergleich starten", type="primary", disabled=(fd_file is None or fw_file is None))

if run:
    fd_df = load_file(fd_file)
    fw_df = load_file(fw_file)

    filter_sap = []
    if sap_input.strip():
        filter_sap = [s.strip() for s in sap_input.strip().splitlines() if s.strip()]

    all_sap = sorted(set(fd_df["SAP"]) | set(fw_df["SAP"]))
    if filter_sap:
        all_sap = [s for s in all_sap if s in filter_sap]
        not_found = [s for s in filter_sap if s not in (set(fd_df["SAP"]) | set(fw_df["SAP"]))]
        if not_found:
            st.warning(f"Nicht gefunden: {', '.join(not_found)}")

    rows = []
    for sap in all_sap:
        fd_rows = fd_df[fd_df["SAP"] == sap]
        fw_rows = fw_df[fw_df["SAP"] == sap]

        name = ""
        if not fd_rows.empty:
            name = fd_rows["Name"].iloc[0]
        elif not fw_rows.empty:
            name = fw_rows["Name"].iloc[0]

        fd_tage = set(fd_rows["Liefertag"].dropna().astype(int))
        fw_tage = set(fw_rows["Liefertag"].dropna().astype(int))

        nur_fd = fd_tage - fw_tage
        nur_fw = fw_tage - fd_tage
        beide = fd_tage & fw_tage

        rows.append({
            "SAP": sap,
            "Name": name,
            "FD Tage": tage_str(fd_tage) if fd_tage else "–",
            "FW Tage": tage_str(fw_tage) if fw_tage else "–",
            "Nur in FD": tage_str(nur_fd) if nur_fd else "–",
            "Nur in FW": tage_str(nur_fw) if nur_fw else "–",
            "In beiden": tage_str(beide) if beide else "–",
            "Unterschied": "JA" if (nur_fd or nur_fw) else "nein",
            "_nur_fd": nur_fd,
            "_nur_fw": nur_fw,
        })

    result_df = pd.DataFrame(rows)

    # Statistik
    total = len(result_df)
    mit_diff = (result_df["Unterschied"] == "JA").sum()
    nur_in_fd = (result_df["FW Tage"] == "–").sum()
    nur_in_fw = (result_df["FD Tage"] == "–").sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Kunden gesamt", total)
    c2.metric("Mit Unterschied", mit_diff)
    c3.metric("Nur in FD", nur_in_fd)
    c4.metric("Nur in FW", nur_in_fw)

    st.divider()

    # Anzeige
    show_cols = ["SAP", "Name", "FD Tage", "FW Tage", "Nur in FD", "Nur in FW", "In beiden", "Unterschied"]

    filter_mode = st.radio("Anzeigen:", ["Alle", "Nur mit Unterschied"], horizontal=True)
    view_df = result_df[show_cols]
    if filter_mode == "Nur mit Unterschied":
        view_df = view_df[result_df["Unterschied"] == "JA"]

    def highlight(row):
        if row["Unterschied"] == "JA":
            return ["background-color: #fff3cd"] * len(row)
        return [""] * len(row)

    st.dataframe(
        view_df.style.apply(highlight, axis=1),
        use_container_width=True,
        hide_index=True,
        height=500,
    )

    # Excel Export
    def build_excel(df_rows):
        wb = Workbook()
        ws = wb.active
        ws.title = "Vergleich"

        header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        header_fill = PatternFill("solid", start_color="2B579A")
        diff_fill = PatternFill("solid", start_color="FFF3CD")
        red_fill = PatternFill("solid", start_color="F8D7DA")
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        thin = Side(style="thin", color="CCCCCC")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        headers = ["SAP", "Name", "FD Liefertage", "FW Liefertage",
                   "Nur in FD", "Nur in FW", "In beiden", "Unterschied"]
        col_widths = [12, 35, 22, 22, 18, 18, 18, 12]

        for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
            cell = ws.cell(row=1, column=ci, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center
            cell.border = border
            ws.column_dimensions[get_column_letter(ci)].width = w

        ws.row_dimensions[1].height = 22

        for ri, row in enumerate(df_rows, 2):
            has_diff = row["Unterschied"] == "JA"
            has_missing = row["FD Tage"] == "–" or row["FW Tage"] == "–"
            fill = red_fill if has_missing else (diff_fill if has_diff else None)

            values = [
                row["SAP"], row["Name"], row["FD Tage"], row["FW Tage"],
                row["Nur in FD"], row["Nur in FW"], row["In beiden"], row["Unterschied"]
            ]
            for ci, val in enumerate(values, 1):
                cell = ws.cell(row=ri, column=ci, value=val)
                cell.font = Font(name="Arial", size=9)
                cell.border = border
                cell.alignment = center if ci != 2 else left
                if fill:
                    cell.fill = fill

            ws.row_dimensions[ri].height = 16

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:H1"

        # 2. Sheet: Nur Unterschiede
        ws2 = wb.create_sheet("Nur Unterschiede")
        for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
            cell = ws2.cell(row=1, column=ci, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center
            cell.border = border
            ws2.column_dimensions[get_column_letter(ci)].width = w
        ws2.row_dimensions[1].height = 22

        ri2 = 2
        for row in df_rows:
            if row["Unterschied"] != "JA":
                continue
            has_missing = row["FD Tage"] == "–" or row["FW Tage"] == "–"
            fill = red_fill if has_missing else diff_fill
            values = [
                row["SAP"], row["Name"], row["FD Tage"], row["FW Tage"],
                row["Nur in FD"], row["Nur in FW"], row["In beiden"], row["Unterschied"]
            ]
            for ci, val in enumerate(values, 1):
                cell = ws2.cell(row=ri2, column=ci, value=val)
                cell.font = Font(name="Arial", size=9)
                cell.border = border
                cell.alignment = center if ci != 2 else left
                cell.fill = fill
            ws2.row_dimensions[ri2].height = 16
            ri2 += 1

        ws2.freeze_panes = "A2"
        ws2.auto_filter.ref = f"A1:H1"

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    excel_buf = build_excel(rows)
    st.download_button(
        "📥 Excel herunterladen",
        data=excel_buf,
        file_name="FD_FW_Vergleich.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

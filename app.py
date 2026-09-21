import io
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Stock Opname Worksheet Generator", page_icon="📊", layout="wide"
)

st.title("📊 Stock Opname Worksheet Generator")
st.write(
    "Aplikasi pemroses & perapi dokumen Excel Stock Opname secara otomatis."
)

st.divider()

# ---------------------------------------------------------
# STEP 1: UPLOAD DOKUMEN
# ---------------------------------------------------------
st.header("1. Upload Dokumen")

col1, col2, col3 = st.columns(3)

with col1:
    template_file = st.file_uploader(
        "Upload Template Master (.xlsx)",
        type=["xlsx"],
        help="Dokumen master/template yang berisi sheet Worksheet, Summary, dll.",
    )

with col2:
    csv_file = st.file_uploader(
        "Upload Data Mentah Inventory (.csv)",
        type=["csv"],
        help="Report inventory baru (bisa separator koma atau titik koma).",
    )

with col3:
    prev_so_file = st.file_uploader(
        "Upload Dokumen Prev SO Referensi (.xlsx)",
        type=["xlsx"],
        help="File SO periode sebelumnya untuk VLOOKUP Batch -> Prev SO & Prev Status.",
    )

st.divider()

# ---------------------------------------------------------
# STEP 2: INPUT METADATA & SUMMARY
# ---------------------------------------------------------
st.header("2. Input Informasi Metadata & Summary")

col_m1, col_m2, col_m3 = st.columns(3)
with col_m1:
    station_input = st.text_input("Station (Cell C3)", value="KNO")
with col_m2:
    location_input = st.text_input(
        "Location (Cell C4)", value="K7, K65 & K42"
    )
with col_m3:
    periode_input = st.text_input(
        "Periode Date (Cell C5)", value="02 - 05 Juni 2026"
    )

st.subheader("Pengaturan Sheet Summary (Dinamis Per Lokasi)")
st.caption(
    "Masukkan detail divisi, PIC, kode lokasi, dan deskripsi lokasi. Anda bisa menambah atau menghapus lokasi sesuai kebutuhan."
)

if "summary_rows" not in st.session_state:
    st.session_state.summary_rows = [
        {
            "division": "LINE MAINTENANCE",
            "pic": "THOMAS",
            "loc_code": "K7",
            "loc_desc": "TOOLROOM KNO",
        },
        {
            "division": "LINE MAINTENANCE",
            "pic": "RAFI",
            "loc_code": "K65",
            "loc_desc": "STORE KNO",
        },
    ]

# Tombol Tambah / Hapus Lokasi
col_b1, col_b2, _ = st.columns([1, 1, 4])
with col_b1:
    if st.button("➕ Tambah Lokasi"):
        st.session_state.summary_rows.append(
            {"division": "", "pic": "", "loc_code": "", "loc_desc": ""}
        )
        st.rerun()

with col_b2:
    if (
        st.button("➖ Hapus Lokasi Terakhir")
        and len(st.session_state.summary_rows) > 1
    ):
        st.session_state.summary_rows.pop()
        st.rerun()

# Form dinamis untuk input summary
summary_data_inputs = []
for idx, row in enumerate(st.session_state.summary_rows):
    st.markdown(f"**Lokasi #{idx + 1}**")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        div = st.text_input(
            f"Division #{idx+1}", value=row["division"], key=f"div_{idx}"
        )
    with c2:
        pic = st.text_input(
            f"PIC #{idx+1}", value=row["pic"], key=f"pic_{idx}"
        )
    with c3:
        loc = st.text_input(
            f"LOC Code #{idx+1}", value=row["loc_code"], key=f"loc_{idx}"
        )
    with c4:
        desc = st.text_input(
            f"LOC Description #{idx+1}",
            value=row["loc_desc"],
            key=f"desc_{idx}",
        )
    summary_data_inputs.append(
        {"division": div, "pic": pic, "loc_code": loc, "loc_desc": desc}
    )

st.divider()

# ---------------------------------------------------------
# STEP 3: PROSES PEMPROSESAN DATA
# ---------------------------------------------------------
st.header("3. Eksekusi & Pemrosesan")


def load_csv_data(uploaded_csv):
    """Membaca CSV baik dengan pemisah koma (,) atau titik koma (;)."""
    try:
        df = pd.read_csv(uploaded_csv, sep=";")
        if df.shape[1] <= 1:
            uploaded_csv.seek(0)
            df = pd.read_csv(uploaded_csv, sep=",")
    except Exception:
        uploaded_csv.seek(0)
        df = pd.read_csv(uploaded_csv, sep=",")
    return df


def build_prev_so_dict(prev_so_file_obj):
    """Membaca file Prev SO dan membuat kamus lookup berdasarkan Batch Key."""
    prev_wb = openpyxl.load_workbook(prev_so_file_obj, data_only=True)
    prev_dict = {}

    if "Worksheet" in prev_wb.sheetnames:
        ws_prev = prev_wb["Worksheet"]

        # Mencari letak header 'Batch', 'Result', 'Status' di sheet Prev SO
        batch_col, result_col, status_col = None, None, None

        # Cek baris 5 sd 7 untuk menemukan header
        for r in range(1, 10):
            for c in range(1, 40):
                val = str(ws_prev.cell(r, c).value or "").strip().lower()
                if val == "batch":
                    batch_col = c
                elif val == "result":
                    result_col = c
                elif val == "status":
                    status_col = c

            if batch_col and result_col and status_col:
                header_row = r
                break

        if batch_col:
            # Mengambil data dari baris setelah header
            start_r = (header_row + 1) if "header_row" in locals() else 8
            for r in range(start_r, ws_prev.max_row + 1):
                batch_val = ws_prev.cell(r, batch_col).value
                if batch_val is not None and str(batch_val).strip() != "":
                    batch_key = str(batch_val).strip()
                    res_val = (
                        ws_prev.cell(r, result_col).value
                        if result_col
                        else None
                    )
                    stat_val = (
                        ws_prev.cell(r, status_col).value
                        if status_col
                        else None
                    )
                    prev_dict[batch_key] = {
                        "prev_so": res_val,
                        "prev_status": stat_val,
                    }
    return prev_dict


if st.button("🚀 Process & Generate Template", type="primary"):
    if not template_file or not csv_file or not prev_so_file:
        st.error(
            "⚠️ Harap upload KETIGA dokumen terlebih dahulu sebelum memproses!"
        )
    else:
        with st.spinner("Sedang memproses data dan merapikan Excel..."):
            # 1. Load Workbook & CSV
            wb = openpyxl.load_workbook(template_file)
            df_raw = load_csv_data(csv_file)
            prev_so_map = build_prev_so_dict(prev_so_file)

            # 2. Hapus Sheet1 dan Sheet2
            for sname in ["Sheet1", "Sheet2", "sheet1", "sheet2"]:
                if sname in wb.sheetnames:
                    del wb[sname]

            # 3. Update Sheet Worksheet
            if "Worksheet" in wb.sheetnames:
                ws = wb["Worksheet"]

                # Update Metadata
                ws["C3"] = f": {station_input}"
                ws["C4"] = f": {location_input}"
                ws["C5"] = f": {periode_input}"

                # Hapus data lama mulai baris 8 ke bawah
                max_r = ws.max_row
                if max_r >= 8:
                    ws.delete_rows(8, amount=max_r - 7 + 1)

                # Pemetaan kolom CSV ke kolom Excel Worksheet (Baris 8+)
                # Kolom A=1, B=2, dst.
                for idx, row_data in df_raw.iterrows():
                    row_idx = 8 + idx

                    # Data dari CSV (menyesuaikan nama header standar CSV report inventory)
                    batch_num = (
                        str(
                            row_data.get("Batch")
                            or row_data.get("BATCH")
                            or ""
                        )
                        .strip()
                    )

                    ws.cell(row_idx, 1, idx + 1)  # No
                    ws.cell(
                        row_idx,
                        2,
                        row_data.get("Count No")
                        or row_data.get("COUNT_NO")
                        or "",
                    )  # Count No
                    ws.cell(
                        row_idx,
                        3,
                        row_data.get("LOC") or row_data.get("LOCATION") or "",
                    )  # LOC
                    ws.cell(
                        row_idx, 4, row_data.get("BIN") or ""
                    )  # BIN
                    ws.cell(
                        row_idx,
                        5,
                        row_data.get("GRB") or row_data.get("GRB_NO") or "",
                    )  # GRB
                    ws.cell(row_idx, 6, batch_num)  # Batch
                    ws.cell(
                        row_idx,
                        7,
                        row_data.get("PN") or row_data.get("PART_NO") or "",
                    )  # PN
                    ws.cell(
                        row_idx,
                        8,
                        row_data.get("SN") or row_data.get("SERIAL_NO") or "",
                    )  # SN
                    ws.cell(
                        row_idx,
                        9,
                        row_data.get("PN Description")
                        or row_data.get("DESCRIPTION")
                        or "",
                    )  # PN Description
                    ws.cell(
                        row_idx,
                        10,
                        row_data.get("QTY Available")
                        or row_data.get("QTY_AVAIL")
                        or 0,
                    )  # QTY Available
                    ws.cell(
                        row_idx,
                        11,
                        row_data.get("QTY Reserved")
                        or row_data.get("QTY_RESV")
                        or 0,
                    )  # QTY Reserved
                    ws.cell(
                        row_idx,
                        12,
                        row_data.get("QTY In Transfer")
                        or row_data.get("QTY_TRANS")
                        or 0,
                    )  # QTY In Transfer
                    ws.cell(
                        row_idx,
                        13,
                        row_data.get("QTY Pending R/I")
                        or row_data.get("QTY_PENDING")
                        or 0,
                    )  # QTY Pending R/I
                    ws.cell(
                        row_idx,
                        14,
                        row_data.get("QTY US")
                        or row_data.get("QTY_US")
                        or 0,
                    )  # QTY US
                    ws.cell(
                        row_idx,
                        15,
                        row_data.get("QTY In Repair")
                        or row_data.get("QTY_REPAIR")
                        or 0,
                    )  # QTY In Repair
                    ws.cell(
                        row_idx,
                        16,
                        row_data.get("Shelf Life Exp") or "",
                    )  # Shelf Life Exp
                    ws.cell(
                        row_idx,
                        17,
                        row_data.get("Condition") or "SV",
                    )  # Condition
                    ws.cell(
                        row_idx,
                        18,
                        row_data.get("Category") or "",
                    )  # Category
                    ws.cell(
                        row_idx,
                        19,
                        row_data.get("Owner") or "",
                    )  # Owner
                    ws.cell(
                        row_idx,
                        20,
                        row_data.get("UOM") or "EA",
                    )  # UOM

                    # --- FORMULA EXCEL DINAMIS ---
                    # Col U (21): Qty eMRO
                    ws.cell(
                        row_idx,
                        21,
                        f"=J{row_idx}+K{row_idx}+M{row_idx}+N{row_idx}",
                    )
                    # Col V (22): Qty Actual
                    ws.cell(row_idx, 22, f"=U{row_idx}")
                    # Col W (23): Diff
                    ws.cell(row_idx, 23, f"=V{row_idx}-U{row_idx}")
                    # Col X (24): Result
                    ws.cell(
                        row_idx,
                        24,
                        f'=IF(U{row_idx}=0,"BUG EMRO??/MISSING??",IF(V{row_idx}=0,"NOT FOUND",IF(V{row_idx}>U{row_idx},"SURPLUS",IF(V{row_idx}<U{row_idx},"MINUS","MATCHED"))))',
                    )
                    # Col Y (25): Status
                    ws.cell(
                        row_idx,
                        25,
                        f'=IF(X{row_idx}="MATCHED","MATCHED","OPEN")',
                    )

                    # Col Z (26) Date, AA (27) Auditor, AB (28) Remark, AC (29) Penyelesaian, AD (30) Corrective Action -> Dikosongkan
                    ws.cell(row_idx, 26, None)
                    ws.cell(row_idx, 27, None)
                    ws.cell(row_idx, 28, None)
                    ws.cell(row_idx, 29, None)
                    ws.cell(row_idx, 30, None)

                    # --- PENUKARAN KOLOM PREV SO & PREV STATUS ---
                    # Col AE (31): Prev SO (Hasil VLOOKUP Result Prev SO)
                    # Col AG (33): Prev Status (Hasil VLOOKUP Status Prev SO)
                    prev_info = prev_so_map.get(
                        batch_num, {"prev_so": None, "prev_status": None}
                    )
                    ws.cell(row_idx, 31, prev_info["prev_so"])  # AE = Prev SO
                    ws.cell(
                        row_idx,
                        32,
                        row_data.get("CAT") or row_data.get("Category") or "",
                    )  # AF = CAT
                    ws.cell(
                        row_idx, 33, prev_info["prev_status"]
                    )  # AG = Prev Status
                    ws.cell(row_idx, 34, None)  # AH = Reason (Dikosongkan)

            # 4. Update Sheet Summary
            if "Summary" in wb.sheetnames:
                ws_sum = wb["Summary"]
                ws_sum["D2"] = f": {station_input}"
                ws_sum["D3"] = f": {periode_input}"

                # Masukkan data baris lokasi
                start_sum_row = 9
                for i, sdata in enumerate(summary_data_inputs):
                    r_curr = start_sum_row + i
                    ws_sum.cell(r_curr, 2, i + 1)  # NO
                    ws_sum.cell(r_curr, 3, sdata["division"])  # DIVISION
                    ws_sum.cell(r_curr, 4, sdata["pic"])  # PIC
                    ws_sum.cell(r_curr, 5, sdata["loc_code"])  # LOC CODE
                    ws_sum.cell(
                        r_curr, 6, sdata["loc_desc"]
                    )  # LOCATION DESCRIPTION

            # Simpan output ke memory buffer
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)

            st.success("✅ Dokumen berhasil diproses dan dirapikan!")

            st.download_button(
                label="📥 Download Hasil Excel Stock Opname",
                data=output_buffer,
                file_name=f"Worksheet_Stock_Opname_{station_input}_Updated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

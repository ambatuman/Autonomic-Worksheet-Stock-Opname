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
        "Upload Data Mentah Inventory (.csv / .xlsx)",
        type=["csv", "xlsx", "xls"],
        help="Report inventory baru (bisa format CSV atau Excel).",
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
# STEP 3: HELPER FUNCTIONS & EXEKUSI
# ---------------------------------------------------------
st.header("3. Eksekusi & Pemrosesan")


def load_raw_inventory_data(uploaded_file_obj):
    """Membaca file data mentah baik berupa Excel (.xlsx/.xls) maupun CSV (pembatas ; atau ,)."""
    fname = uploaded_file_obj.name.lower()
    if fname.endswith(".xlsx") or fname.endswith(".xls"):
        df = pd.read_excel(uploaded_file_obj)
    else:
        try:
            df = pd.read_csv(uploaded_file_obj, sep=";")
            if df.shape[1] <= 1:
                uploaded_file_obj.seek(0)
                df = pd.read_csv(uploaded_file_obj, sep=",")
        except Exception:
            uploaded_file_obj.seek(0)
            df = pd.read_csv(uploaded_file_obj, sep=",")

    # Clean & normalize column names to UPPERCASE
    df.columns = df.columns.astype(str).str.strip().str.upper()
    return df


def get_val(row, *possible_keys, default=None):
    """Helper untuk mengambil data secara Case-Insensitive & fleksibel dari dictionary baris."""
    for key in possible_keys:
        k_upper = key.strip().upper()
        if k_upper in row and pd.notna(row[k_upper]):
            val = row[k_upper]
            # Jika numerik, coba kembalikan int/float, bukan NaN
            return val
    return default


def build_prev_so_dict(prev_so_file_obj):
    """Membaca file Prev SO dan membuat kamus lookup berdasarkan Batch Key."""
    prev_wb = openpyxl.load_workbook(prev_so_file_obj, data_only=True)
    prev_dict = {}

    if "Worksheet" in prev_wb.sheetnames:
        ws_prev = prev_wb["Worksheet"]

        batch_col, result_col, status_col = None, None, None
        header_row = 7

        # Mencari letak header 'Batch', 'Result', 'Status' di sheet Prev SO
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
            for r in range(header_row + 1, ws_prev.max_row + 1):
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
            wb = openpyxl.load_workbook(template_file)
            df_raw = load_raw_inventory_data(csv_file)
            prev_so_map = build_prev_so_dict(prev_so_file)

            # Hapus Sheet1 & Sheet2 jika ada
            for sname in ["Sheet1", "Sheet2", "sheet1", "sheet2"]:
                if sname in wb.sheetnames:
                    del wb[sname]

            # Update Sheet Worksheet
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

                # Isi data mentah baru mulai dari baris ke-8
                for idx, row_data in df_raw.iterrows():
                    row_idx = 8 + idx

                    batch_num = str(
                        get_val(
                            row_data, "BATCH", "BATCH NO", "BATCH_NO", default=""
                        )
                    ).strip()

                    ws.cell(row_idx, 1, idx + 1)  # A: No
                    ws.cell(
                        row_idx,
                        2,
                        get_val(row_data, "COUNT NO", "COUNT_NO", default=""),
                    )  # B: Count No
                    ws.cell(
                        row_idx,
                        3,
                        get_val(row_data, "LOCATION", "LOC", default=""),
                    )  # C: LOC
                    ws.cell(
                        row_idx, 4, get_val(row_data, "BIN", default="")
                    )  # D: BIN
                    ws.cell(
                        row_idx,
                        5,
                        get_val(row_data, "GRB", "GRB NO", "GRB_NO", default=""),
                    )  # E: GRB
                    ws.cell(row_idx, 6, batch_num)  # F: Batch
                    ws.cell(
                        row_idx,
                        7,
                        get_val(
                            row_data, "PN", "PART NO", "PART_NO", default=""
                        ),
                    )  # G: PN
                    ws.cell(
                        row_idx,
                        8,
                        get_val(
                            row_data, "SN", "SERIAL NO", "SERIAL_NO", default=""
                        ),
                    )  # H: SN
                    ws.cell(
                        row_idx,
                        9,
                        get_val(
                            row_data,
                            "PN DESCRIPTION",
                            "DESCRIPTION",
                            "PN DESC",
                            default="",
                        ),
                    )  # I: PN Description

                    # QTY Columns (10-15)
                    ws.cell(
                        row_idx,
                        10,
                        get_val(
                            row_data,
                            "QTY AVAILABLE",
                            "QTY_AVAIL",
                            "QTY AVAIL",
                            default=0,
                        ),
                    )  # J
                    ws.cell(
                        row_idx,
                        11,
                        get_val(
                            row_data,
                            "QTY RESERVED",
                            "QTY_RESV",
                            "QTY RESV",
                            default=0,
                        ),
                    )  # K
                    ws.cell(
                        row_idx,
                        12,
                        get_val(
                            row_data,
                            "QTY IN TRANSFER",
                            "QTY_TRANS",
                            "QTY TRANS",
                            default=0,
                        ),
                    )  # L
                    ws.cell(
                        row_idx,
                        13,
                        get_val(
                            row_data,
                            "QTY PENDING RI",
                            "QTY PENDING R/I",
                            "QTY_PENDING",
                            default=0,
                        ),
                    )  # M
                    ws.cell(
                        row_idx,
                        14,
                        get_val(row_data, "QTY US", "QTY_US", default=0),
                    )  # N
                    ws.cell(
                        row_idx,
                        15,
                        get_val(
                            row_data,
                            "QTY IN REPAIR",
                            "QTY_REPAIR",
                            "QTY REPAIR",
                            default=0,
                        ),
                    )  # O

                    ws.cell(
                        row_idx,
                        16,
                        get_val(
                            row_data,
                            "SHELF LIFE EXPIRATION",
                            "SHELF LIFE EXP",
                            "TOOL LIFE EXPIRATION",
                            default="",
                        ),
                    )  # P: Shelf Life Exp
                    ws.cell(
                        row_idx,
                        17,
                        get_val(row_data, "CONDITION", default="SV"),
                    )  # Q: Condition
                    ws.cell(
                        row_idx, 18, get_val(row_data, "CATEGORY", default="")
                    )  # R: Category
                    ws.cell(
                        row_idx, 19, get_val(row_data, "OWNER", default="")
                    )  # S: Owner
                    ws.cell(
                        row_idx, 20, get_val(row_data, "UOM", default="EA")
                    )  # T: UOM

                    # --- FORMULA EXCEL DINAMIS ---
                    ws.cell(
                        row_idx,
                        21,
                        f"=J{row_idx}+K{row_idx}+M{row_idx}+N{row_idx}",
                    )  # U: Qty eMRO
                    ws.cell(row_idx, 22, f"=U{row_idx}")  # V: Qty Actual
                    ws.cell(
                        row_idx, 23, f"=V{row_idx}-U{row_idx}"
                    )  # W: Diff
                    ws.cell(
                        row_idx,
                        24,
                        f'=IF(U{row_idx}=0,"BUG EMRO??/MISSING??",IF(V{row_idx}=0,"NOT FOUND",IF(V{row_idx}>U{row_idx},"SURPLUS",IF(V{row_idx}<U{row_idx},"MINUS","MATCHED"))))',
                    )  # X: Result
                    ws.cell(
                        row_idx,
                        25,
                        f'=IF(X{row_idx}="MATCHED","MATCHED","OPEN")',
                    )  # Y: Status

                    # Dikosongkan
                    ws.cell(row_idx, 26, None)  # Z: Date
                    ws.cell(row_idx, 27, None)  # AA: Auditor
                    ws.cell(row_idx, 28, None)  # AB: Remark
                    ws.cell(row_idx, 29, None)  # AC: Penyelesaian
                    ws.cell(row_idx, 30, None)  # AD: Corrective Action

                    # --- PENUKARAN KOLOM PREV SO & PREV STATUS ---
                    prev_info = prev_so_map.get(
                        batch_num, {"prev_so": None, "prev_status": None}
                    )
                    ws.cell(
                        row_idx, 31, prev_info["prev_so"]
                    )  # AE: Prev SO
                    ws.cell(
                        row_idx,
                        32,
                        get_val(row_data, "CAT", "CATEGORY", default=""),
                    )  # AF: CAT
                    ws.cell(
                        row_idx, 33, prev_info["prev_status"]
                    )  # AG: Prev Status
                    ws.cell(row_idx, 34, None)  # AH: Reason (Dikosongkan)

            # Update Sheet Summary
            if "Summary" in wb.sheetnames:
                ws_sum = wb["Summary"]
                ws_sum["D2"] = f": {station_input}"
                ws_sum["D3"] = f": {periode_input}"

                start_sum_row = 9
                for i, sdata in enumerate(summary_data_inputs):
                    r_curr = start_sum_row + i
                    ws_sum.cell(r_curr, 2, i + 1)
                    ws_sum.cell(r_curr, 3, sdata["division"])
                    ws_sum.cell(r_curr, 4, sdata["pic"])
                    ws_sum.cell(r_curr, 5, sdata["loc_code"])
                    ws_sum.cell(r_curr, 6, sdata["loc_desc"])

            # Save ke memory buffer untuk didownload
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

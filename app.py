import io
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Stock Opname Worksheet Generator", page_icon="📊", layout="wide"
)

st.title("📊 Stock Opname Worksheet Generator")
st.write(
    "Aplikasi pemroses & otomatisasi laporan Excel Stock Opname secara presisi."
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
    station_input = st.text_input("Station", value="SUB")
with col_m2:
    location_input = st.text_input(
        "Location", value="S1, K86, K88, K87, K56 & K58"
    )
with col_m3:
    periode_input = st.text_input(
        "Periode Date", value="SEPTEMBER 2026"
    )

st.subheader("Pengaturan Sheet Summary (Dinamis Per Lokasi)")
st.caption(
    "Masukkan detail divisi, PIC, kode lokasi, dan deskripsi lokasi."
)

if "summary_rows" not in st.session_state:
    st.session_state.summary_rows = [
        {
            "division": "LINE MAINTENANCE",
            "pic": "THOMAS",
            "loc_code": "K86",
            "loc_desc": "STORE SUB",
        }
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


def clean_batch_str(val):
    """Pembersih penulisan string Batch agar match 100%."""
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def load_raw_inventory_data(uploaded_file_obj):
    """Membaca file data mentah inventory & melakukan sorting lokasi/BIN."""
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

    df.columns = df.columns.astype(str).str.strip().str.upper()

    # --- FITUR SORTING ---
    sort_cols = []
    for col in ["LOCATION", "LOC", "BIN", "PART NO", "PN", "BATCH", "BATCH NO"]:
        if col in df.columns and col not in sort_cols:
            sort_cols.append(col)
    
    if sort_cols:
        df = df.sort_values(by=sort_cols, ascending=True).reset_index(drop=True)

    return df


def get_val(row, *possible_keys, default=None):
    """Mencari nilai kolom secara case-insensitive & aman dari NaN."""
    for key in possible_keys:
        k_upper = key.strip().upper()
        if k_upper in row and pd.notna(row[k_upper]):
            val = row[k_upper]
            return val
    return default


def set_cell_safe(ws, row, col, value):
    """Mengisi sel tanpa merusak jika menimpa MergedCell."""
    cell = ws.cell(row=row, column=col)
    if type(cell).__name__ == "MergedCell":
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
                break
    else:
        cell.value = value


def build_prev_so_dict(prev_so_file_obj):
    """Pindai otomatis header Prev SO, toleran terhadap variasi nama kolom."""
    try:
        df_raw = pd.read_excel(prev_so_file_obj, sheet_name="Worksheet", header=None)
    except Exception:
        prev_so_file_obj.seek(0)
        df_raw = pd.read_excel(prev_so_file_obj, header=None)

    # Deteksi baris header secara fleksibel
    header_idx = None
    for idx, r in df_raw.iterrows():
        row_str_vals = [str(v).strip().upper() for v in r.values if pd.notna(v)]
        if "BATCH" in row_str_vals:
            header_idx = idx
            break

    prev_dict = {}
    if header_idx is not None:
        prev_so_file_obj.seek(0)
        try:
            df_prev = pd.read_excel(prev_so_file_obj, sheet_name="Worksheet", header=header_idx)
        except Exception:
            prev_so_file_obj.seek(0)
            df_prev = pd.read_excel(prev_so_file_obj, header=header_idx)

        df_prev.columns = [str(c).strip().upper() for c in df_prev.columns]

        # Identifikasi kolom target untuk Prev SO & Prev Status
        batch_col = next((c for c in df_prev.columns if "BATCH" in c), None)
        so_col = next((c for c in df_prev.columns if any(k in c for k in ["RESULT", "PREV SO", "SO", "RESULT PREV"])), None)
        status_col = next((c for c in df_prev.columns if any(k in c for k in ["STATUS", "PREV STATUS", "STATUS PREV"])), None)

        if batch_col:
            for _, row in df_prev.iterrows():
                batch_key = clean_batch_str(row.get(batch_col))
                if batch_key:
                    res_val = row.get(so_col) if so_col and pd.notna(row.get(so_col)) else None
                    stat_val = row.get(status_col) if status_col and pd.notna(row.get(status_col)) else None

                    prev_dict[batch_key] = {
                        "prev_so": res_val,
                        "prev_status": stat_val,
                    }
    return prev_dict


if st.button("🚀 Process & Generate Template", type="primary"):
    if not template_file or not csv_file or not prev_so_file:
        st.error("⚠️ Harap upload KETIGA dokumen terlebih dahulu!")
    else:
        with st.spinner("Sedang memproses data dan merapikan Excel..."):
            wb = openpyxl.load_workbook(template_file)
            df_raw = load_raw_inventory_data(csv_file)
            prev_so_map = build_prev_so_dict(prev_so_file)

            # Hapus Sheet1 & Sheet2 jika ada
            for sname in ["Sheet1", "Sheet2", "sheet1", "sheet2"]:
                if sname in wb.sheetnames:
                    del wb[sname]

            # -----------------------------------------------------
            # UPDATE SHEET WORKSHEET
            # -----------------------------------------------------
            if "Worksheet" in wb.sheetnames:
                ws = wb["Worksheet"]

                # Update Metadata Header
                set_cell_safe(ws, 3, 3, f": {station_input}")
                set_cell_safe(ws, 4, 3, f": {location_input}")
                set_cell_safe(ws, 5, 3, f": {periode_input}")

                # Hapus isi baris lama dari baris 8 ke bawah
                max_r = ws.max_row
                if max_r >= 8:
                    ws.delete_rows(8, amount=max_r - 7 + 1)

                # Populasi Data Mentah Baru
                for idx, row_data in df_raw.iterrows():
                    row_idx = 8 + idx

                    raw_batch = get_val(row_data, "BATCH", "BATCH NO", "BATCH_NO", default="")
                    batch_num = clean_batch_str(raw_batch)

                    set_cell_safe(ws, row_idx, 1, idx + 1)  # A: No
                    set_cell_safe(ws, row_idx, 2, get_val(row_data, "COUNT NO", "COUNT_NO", default=""))  # B: Count No
                    set_cell_safe(ws, row_idx, 3, get_val(row_data, "LOCATION", "LOC", default=""))  # C: LOC
                    set_cell_safe(ws, row_idx, 4, get_val(row_data, "BIN", default=""))  # D: BIN
                    set_cell_safe(ws, row_idx, 5, get_val(row_data, "GRB", "GRB NO", "GRB_NO", default=""))  # E: GRB
                    set_cell_safe(ws, row_idx, 6, batch_num)  # F: Batch
                    set_cell_safe(ws, row_idx, 7, get_val(row_data, "PN", "PART NO", "PART_NO", default=""))  # G: PN
                    set_cell_safe(ws, row_idx, 8, get_val(row_data, "SN", "SERIAL NO", "SERIAL_NO", default=""))  # H: SN
                    set_cell_safe(ws, row_idx, 9, get_val(row_data, "PN DESCRIPTION", "DESCRIPTION", "PN DESC", default=""))  # I: PN Description

                    # QTY Columns (10-15)
                    set_cell_safe(ws, row_idx, 10, get_val(row_data, "QTY AVAILABLE", "QTY_AVAIL", "QTY AVAIL", default=0))  # J: Available
                    set_cell_safe(ws, row_idx, 11, get_val(row_data, "QTY RESERVED", "QTY_RESV", "QTY RESV", default=0))  # K: Reserved
                    set_cell_safe(ws, row_idx, 12, get_val(row_data, "QTY IN TRANSFER", "QTY_TRANS", "QTY TRANS", default=0))  # L: Transfer
                    set_cell_safe(ws, row_idx, 13, get_val(row_data, "QTY PENDING RI", "QTY PENDING R/I", "QTY_PENDING", default=0))  # M: Pending RI
                    set_cell_safe(ws, row_idx, 14, get_val(row_data, "QTY US", "QTY_US", default=0))  # N: US
                    set_cell_safe(ws, row_idx, 15, get_val(row_data, "QTY IN REPAIR", "QTY_REPAIR", "QTY REPAIR", default=0))  # O: In Repair

                    set_cell_safe(ws, row_idx, 16, get_val(row_data, "SHELF LIFE EXPIRATION", "SHELF LIFE EXP", "TOOL LIFE EXPIRATION", default=""))  # P
                    set_cell_safe(ws, row_idx, 17, get_val(row_data, "CONDITION", default="SV"))  # Q: Condition
                    set_cell_safe(ws, row_idx, 18, get_val(row_data, "CATEGORY", default=""))  # R: Category
                    set_cell_safe(ws, row_idx, 19, get_val(row_data, "OWNER", default=""))  # S: Owner
                    set_cell_safe(ws, row_idx, 20, get_val(row_data, "UOM", default="EA"))  # T: UOM

                    # --- FORMULA EXCEL DINAMIS ---
                    set_cell_safe(ws, row_idx, 21, f"=J{row_idx}+K{row_idx}+M{row_idx}+N{row_idx}")  # U: Qty eMRO
                    
                    # QTY ACTUAL (KOLOM V) DI-BLANK-KAN UNTUK PENGISIAN MANUAL AUDITOR
                    set_cell_safe(ws, row_idx, 22, None)  # V: Qty Actual
                    
                    set_cell_safe(ws, row_idx, 23, f"=V{row_idx}-U{row_idx}")  # W: Diff
                    set_cell_safe(ws, row_idx, 24, f'=IF(U{row_idx}=0,"BUG EMRO??/MISSING??",IF(V{row_idx}=0,"NOT FOUND",IF(V{row_idx}>U{row_idx},"SURPLUS",IF(V{row_idx}<U{row_idx},"MINUS","MATCHED"))))')  # X: Result
                    set_cell_safe(ws, row_idx, 25, f'=IF(X{row_idx}="MATCHED","MATCHED","OPEN")')  # Y: Status

                    # Kolom manual yang dikosongkan
                    set_cell_safe(ws, row_idx, 26, None)  # Z: Date
                    set_cell_safe(ws, row_idx, 27, None)  # AA: Auditor
                    set_cell_safe(ws, row_idx, 28, None)  # AB: Remark
                    set_cell_safe(ws, row_idx, 29, None)  # AC: Penyelesaian
                    set_cell_safe(ws, row_idx, 30, None)  # AD: Corrective Action

                    # --- PREV SO & PREV STATUS (AE & AG) ---
                    prev_info = prev_so_map.get(batch_num, {"prev_so": None, "prev_status": None})
                    
                    set_cell_safe(ws, row_idx, 31, prev_info["prev_so"])  # AE (31): Prev SO
                    set_cell_safe(ws, row_idx, 32, get_val(row_data, "CAT", "CATEGORY", default=""))  # AF (32): CAT
                    set_cell_safe(ws, row_idx, 33, prev_info["prev_status"])  # AG (33): Prev Status
                    set_cell_safe(ws, row_idx, 34, None)  # AH (34): Reason (Blank)

            # -----------------------------------------------------
            # UPDATE SHEET SUMMARY
            # -----------------------------------------------------
            if "Summary" in wb.sheetnames:
                ws_sum = wb["Summary"]
                set_cell_safe(ws_sum, 3, 4, f": {station_input}")  # D3: Station
                set_cell_safe(ws_sum, 4, 4, f": {periode_input}")  # D4: Periode

                start_sum_row = 11  # Data Summary dimulai di baris 11
                for i, sdata in enumerate(summary_data_inputs):
                    r_curr = start_sum_row + i
                    set_cell_safe(ws_sum, r_curr, 2, i + 1)  # B: NO
                    set_cell_safe(ws_sum, r_curr, 3, sdata["division"])  # C: DIVISION
                    set_cell_safe(ws_sum, r_curr, 4, sdata["pic"])  # D: PIC
                    set_cell_safe(ws_sum, r_curr, 5, sdata["loc_code"])  # E: LOC CODE
                    set_cell_safe(ws_sum, r_curr, 6, sdata["loc_desc"])  # F: LOCATION DESCRIPTION

            # Save ke memory buffer untuk download
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)

            st.success("✅ Otomasi laporan berhasil diproses dengan presisi!")

            st.download_button(
                label="📥 Download Hasil Excel Stock Opname",
                data=output_buffer,
                file_name=f"Worksheet_Stock_Opname_{station_input}_Updated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

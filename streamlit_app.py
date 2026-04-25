import streamlit as st
import pandas as pd
from io import BytesIO
import re

REPORTING_DATE = pd.Timestamp("2025-12-31")


def inject_custom_css():
    st.markdown("""
    <style>
        .main > div {
            padding-top: 1rem;
        }

        .block-container {
            padding-top: 1.2rem;
            padding-bottom: 2rem;
            padding-left: 1.5rem;
            padding-right: 1.5rem;
        }

        .custom-title {
            font-size: 2rem;
            font-weight: 700;
            margin-bottom: 0.2rem;
        }

        .custom-subtitle {
            color: #666;
            font-size: 0.95rem;
            margin-bottom: 1rem;
        }

        .info-box {
            background-color: #f8f9fa;
            border: 1px solid #e9ecef;
            padding: 14px 16px;
            border-radius: 12px;
            margin-bottom: 12px;
        }

        .status-card {
            border-radius: 14px;
            padding: 14px 16px;
            color: white;
            margin-bottom: 10px;
        }

        .status-green {
            background: linear-gradient(135deg, #16a34a, #22c55e);
        }

        .status-yellow {
            background: linear-gradient(135deg, #ca8a04, #eab308);
            color: #111827;
        }

        .status-red {
            background: linear-gradient(135deg, #dc2626, #ef4444);
        }

        .status-blue {
            background: linear-gradient(135deg, #2563eb, #3b82f6);
        }

        .status-label {
            font-size: 0.9rem;
            opacity: 0.95;
            margin-bottom: 4px;
        }

        .status-value {
            font-size: 1.5rem;
            font-weight: 700;
            line-height: 1.2;
        }

        .stDownloadButton button, .stButton button {
            border-radius: 10px !important;
            font-weight: 600 !important;
        }

        div[data-testid="stMetric"] {
            background-color: #ffffff;
            border: 1px solid #e5e7eb;
            padding: 12px;
            border-radius: 12px;
        }

        .search-box-note {
            color: #6b7280;
            font-size: 0.85rem;
            margin-top: -6px;
            margin-bottom: 10px;
        }
    </style>
    """, unsafe_allow_html=True)


def status_card(label, value, color_class):
    st.markdown(
        f"""
        <div class="status-card {color_class}">
            <div class="status-label">{label}</div>
            <div class="status-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def parse_mixed_excel_date(value):
    if pd.isna(value):
        return pd.NaT

    if isinstance(value, pd.Timestamp):
        return value

    text = str(value).strip().replace("\xa0", "").replace("  ", " ")

    if text == "" or text.lower() in ["nan", "none", "nat"]:
        return pd.NaT

    try:
        num = float(text)
        if num > 1000 and text.replace(".", "", 1).isdigit():
            return pd.Timestamp("1899-12-30") + pd.to_timedelta(int(num), unit="D")
    except Exception:
        pass

    if re.match(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}", text):
        return pd.to_datetime(text, errors="coerce", yearfirst=True)

    return pd.to_datetime(text, errors="coerce", dayfirst=True)


def normalize_kode_aset(value):
    if pd.isna(value):
        return pd.NA

    text = str(value).strip().replace("\xa0", "")

    if text == "" or text.lower() in ["nan", "none"]:
        return pd.NA

    try:
        num = float(text)
        if num.is_integer():
            return str(int(num))
    except Exception:
        pass

    return text


def safe_sheet_name(name):
    text = str(name)
    for ch in ["/", "\\", ":", "*", "?", "[", "]"]:
        text = text.replace(ch, "_")
    return text[:31] if text else "Sheet"


def calculate_depreciation_monthly(
    initial_cost,
    acquisition_date,
    useful_life_years,
    reporting_date=REPORTING_DATE,
    capitalizations=None,
    corrections=None
):
    if capitalizations is None:
        capitalizations = []
    if corrections is None:
        corrections = []

    acquisition_date = parse_mixed_excel_date(acquisition_date)
    reporting_date = parse_mixed_excel_date(reporting_date)

    if pd.isna(acquisition_date) or pd.isna(reporting_date):
        return []

    if acquisition_date > reporting_date:
        return []

    original_life_months = int(float(useful_life_years) * 12)
    remaining_life_months = original_life_months

    book_value = float(initial_cost)
    accumulated_dep = 0.0

    cap_dict = {}
    for cap in capitalizations:
        cap_date = parse_mixed_excel_date(cap.get("Tanggal Kapitalisasi"))
        if pd.notna(cap_date) and cap_date <= reporting_date:
            key = (cap_date.year, cap_date.month)
            cap_dict.setdefault(key, []).append(cap)

    corr_dict = {}
    for corr in corrections:
        corr_date = parse_mixed_excel_date(corr.get("Tanggal Koreksi"))
        if pd.notna(corr_date) and corr_date <= reporting_date:
            key = (corr_date.year, corr_date.month)
            corr_dict.setdefault(key, []).append(corr)

    current_year = acquisition_date.year
    current_month = acquisition_date.month
    schedule = []

    while (current_year < reporting_date.year) or (
        current_year == reporting_date.year and current_month <= reporting_date.month
    ):
        current_key = (current_year, current_month)

        kapitalisasi_bulan_ini = 0.0
        koreksi_bulan_ini = 0.0
        tambahan_usia_bulan_ini = 0

        if current_key in cap_dict:
            for cap in cap_dict[current_key]:
                cap_amount = float(cap.get("Jumlah", 0) or 0)
                tambahan_usia_tahun = float(cap.get("Tambahan Usia", 0) or 0)
                tambahan_usia_bulan = int(tambahan_usia_tahun * 12)

                kapitalisasi_bulan_ini += cap_amount
                tambahan_usia_bulan_ini += tambahan_usia_bulan

            book_value += kapitalisasi_bulan_ini
            remaining_life_months = min(
                remaining_life_months + tambahan_usia_bulan_ini,
                original_life_months
            )

        if current_key in corr_dict:
            for corr in corr_dict[current_key]:
                corr_amount = float(corr.get("Jumlah", 0) or 0)
                koreksi_bulan_ini += corr_amount

            book_value = max(book_value - koreksi_bulan_ini, 0)

        monthly_dep = 0.0
        if remaining_life_months > 0 and book_value > 0:
            monthly_dep = book_value / remaining_life_months
            accumulated_dep += monthly_dep
            book_value -= monthly_dep
            remaining_life_months -= 1

        schedule.append({
            "Tahun": current_year,
            "Bulan": current_month,
            "Periode": f"{current_year}-{current_month:02d}",
            "Kapitalisasi Bulan Ini": round(kapitalisasi_bulan_ini, 2),
            "Tambahan Usia Bulan Ini": tambahan_usia_bulan_ini,
            "Koreksi Bulan Ini": round(koreksi_bulan_ini, 2),
            "Penyusutan Bulan Berjalan": round(monthly_dep, 2),
            "Akumulasi Penyusutan": round(accumulated_dep, 2),
            "Nilai Buku Akhir": round(book_value, 2),
            "Sisa Masa Manfaat (Bulan)": remaining_life_months,
            "Sisa Masa Manfaat (Tahun)": round(remaining_life_months / 12, 2),
        })

        current_month += 1
        if current_month > 12:
            current_month = 1
            current_year += 1

    return schedule


def create_template_excel():
    buffer = BytesIO()

    df_assets = pd.DataFrame([
        {
            "Kode Aset": "211506",
            "Harga Perolehan Awal (Rp)": 401150000,
            "Tanggal Perolehan": "14/08/2017",
            "Masa Manfaat (tahun)": 8
        }
    ])

    df_caps = pd.DataFrame([
        {
            "Kode Aset": "211506",
            "Tanggal Kapitalisasi": "12/12/2017",
            "Jumlah": 206990000,
            "Tambahan Usia": 4
        }
    ])

    df_corrs = pd.DataFrame([
        {
            "Kode Aset": "211506",
            "Tanggal Koreksi": "05/10/2025",
            "Jumlah": 2000000
        }
    ])

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df_assets.to_excel(writer, index=False, sheet_name="Data Aset")
        df_caps.to_excel(writer, index=False, sheet_name="Kapitalisasi")
        df_corrs.to_excel(writer, index=False, sheet_name="Koreksi")

        workbook = writer.book
        date_fmt = workbook.add_format({"num_format": "dd/mm/yyyy"})
        money_fmt = workbook.add_format({"num_format": "#,##0.00"})

        ws1 = writer.sheets["Data Aset"]
        ws2 = writer.sheets["Kapitalisasi"]
        ws3 = writer.sheets["Koreksi"]

        ws1.set_column("A:A", 18)
        ws1.set_column("B:B", 24, money_fmt)
        ws1.set_column("C:C", 18, date_fmt)
        ws1.set_column("D:D", 18)

        ws2.set_column("A:A", 18)
        ws2.set_column("B:B", 18, date_fmt)
        ws2.set_column("C:C", 20, money_fmt)
        ws2.set_column("D:D", 18)

        ws3.set_column("A:A", 18)
        ws3.set_column("B:B", 18, date_fmt)
        ws3.set_column("C:C", 20, money_fmt)

    buffer.seek(0)
    return buffer.getvalue()


def convert_df_to_excel_with_sheets(results, schedules, skipped_rows=None, anomaly_rows=None, total_rows=0):
    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        workbook = writer.book
        money_fmt = workbook.add_format({"num_format": "#,##0.00"})
        int_fmt = workbook.add_format({"num_format": "0"})
        bold_fmt = workbook.add_format({"bold": True})

        results_df = pd.DataFrame(results)
        results_df.to_excel(writer, index=False, sheet_name="Ringkasan")

        ws_ringkasan = writer.sheets["Ringkasan"]
        ws_ringkasan.set_column("A:A", 20)
        ws_ringkasan.set_column("B:C", 18)
        ws_ringkasan.set_column("D:F", 22, money_fmt)
        ws_ringkasan.set_column("G:G", 20, int_fmt)

        for asset_code, schedule in schedules.items():
            schedule_df = pd.DataFrame(schedule)
            sheet_name = safe_sheet_name(asset_code)

            schedule_df.to_excel(writer, sheet_name=sheet_name, startrow=2, index=False)
            ws = writer.sheets[sheet_name]

            ws.write(0, 0, "Kode Aset", bold_fmt)
            ws.write(0, 1, asset_code)
            ws.write(1, 0, "Tanggal Pelaporan", bold_fmt)
            ws.write(1, 1, REPORTING_DATE.strftime("%d/%m/%Y"))

            ws.set_column("A:C", 18)
            ws.set_column("D:G", 22, money_fmt)
            ws.set_column("H:I", 22, int_fmt)

        ws_reviu = workbook.add_worksheet("Reviu Hasil")
        writer.sheets["Reviu Hasil"] = ws_reviu

        processed_rows = len(results)
        skipped_count = len(skipped_rows) if skipped_rows else 0
        anomaly_count = len(anomaly_rows) if anomaly_rows else 0

        ws_reviu.write(0, 0, "Ringkasan Reviu", bold_fmt)
        ws_reviu.write(2, 0, "Jumlah total baris", bold_fmt)
        ws_reviu.write(2, 1, total_rows, int_fmt)

        ws_reviu.write(3, 0, "Jumlah baris berhasil diproses", bold_fmt)
        ws_reviu.write(3, 1, processed_rows, int_fmt)

        ws_reviu.write(4, 0, "Jumlah baris dilewati", bold_fmt)
        ws_reviu.write(4, 1, skipped_count, int_fmt)

        ws_reviu.write(5, 0, "Jumlah input aset anomali", bold_fmt)
        ws_reviu.write(5, 1, anomaly_count, int_fmt)

        start_row_skip = 8
        ws_reviu.write(start_row_skip, 0, "Daftar Baris yang Dilewati", bold_fmt)
        skipped_df = pd.DataFrame(
            skipped_rows if skipped_rows else [],
            columns=["Baris Excel", "Kode Aset", "Alasan"]
        )
        skipped_df.to_excel(writer, index=False, sheet_name="Reviu Hasil", startrow=start_row_skip + 1)

        start_row_anom = start_row_skip + 3 + max(len(skipped_df), 1)
        ws_reviu.write(start_row_anom, 0, "Daftar Input Aset Tidak Logis / Anomali", bold_fmt)
        anomaly_df = pd.DataFrame(
            anomaly_rows if anomaly_rows else [],
            columns=["Kode Aset", "Jenis Anomali", "Tanggal Aset", "Tanggal Transaksi", "Keterangan"]
        )
        anomaly_df.to_excel(writer, index=False, sheet_name="Reviu Hasil", startrow=start_row_anom + 1)

        ws_reviu.set_column("A:A", 18)
        ws_reviu.set_column("B:B", 28)
        ws_reviu.set_column("C:D", 18)
        ws_reviu.set_column("E:E", 70)

    buffer.seek(0)
    return buffer.getvalue()


def app():
    st.set_page_config(
        page_title="Depresiasi GL Bulanan",
        page_icon="📉",
        layout="wide"
    )

    inject_custom_css()

    st.markdown('<div class="custom-title">📉 Depresiasi GL Bulanan</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="custom-subtitle">Perhitungan penyusutan bulanan dengan tanggal pelaporan otomatis 31 Desember 2025.</div>',
        unsafe_allow_html=True
    )

    with st.sidebar:
        st.header("⚙️ Panel Aplikasi")
        st.markdown(f"""
        <div class="info-box">
        <b>Tanggal Pelaporan</b><br>
        {REPORTING_DATE.strftime("%d/%m/%Y")}
        </div>
        """, unsafe_allow_html=True)

        st.markdown("### 📥 Template")
        template_file = create_template_excel()
        st.download_button(
            "⬇️ Download Template Excel",
            template_file,
            "template_penyusutan_bulanan_2025.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        st.markdown("### 📤 Upload File")
        uploaded_file = st.file_uploader(
            "Unggah file Excel",
            type=["xlsx"],
            label_visibility="collapsed"
        )

        st.markdown("---")
        st.markdown("### ℹ️ Format Data")
        st.markdown("""
        **Sheet 1 - Data Aset**
        - Kode Aset
        - Harga Perolehan Awal (Rp)
        - Tanggal Perolehan
        - Masa Manfaat (tahun)

        **Sheet 2 - Kapitalisasi**
        - Kode Aset
        - Tanggal Kapitalisasi
        - Jumlah
        - Tambahan Usia

        **Sheet 3 - Koreksi**
        - Kode Aset
        - Tanggal Koreksi
        - Jumlah
        """)

        st.markdown("---")
        st.caption("Tambahan usia diisi dalam tahun dan akan dikonversi ke bulan.")

    with st.expander("📖 Petunjuk Penggunaan", expanded=False):
        st.markdown("""
        1. Unduh template Excel.  
        2. Isi data aset, kapitalisasi, dan koreksi.  
        3. Unggah file Excel pada panel kiri.  
        4. Sistem akan menghitung penyusutan sampai **31/12/2025**.  
        5. Hasil dapat diunduh dalam format Excel.

        **Catatan penting**
        - Sheet kapitalisasi dan koreksi boleh kosong.
        - Kapitalisasi/koreksi sebelum tanggal perolehan induk akan dicatat sebagai **anomali**.
        - Tambahan usia maksimal sampai masa manfaat awal/induk.
        """)

    if uploaded_file is None:
        st.info("Silakan upload file Excel pada panel kiri untuk mulai memproses data.")
        return

    try:
        excel_data = pd.ExcelFile(uploaded_file, engine="openpyxl")
        sheet_names = excel_data.sheet_names

        if len(sheet_names) < 1:
            st.error("File Excel tidak memiliki sheet yang dapat dibaca.")
            return

        assets_df = excel_data.parse(sheet_name=0)

        if len(sheet_names) >= 2:
            capitalizations_df = excel_data.parse(sheet_name=1)
        else:
            capitalizations_df = pd.DataFrame(columns=[
                "Kode Aset", "Tanggal Kapitalisasi", "Jumlah", "Tambahan Usia"
            ])

        if len(sheet_names) >= 3:
            corrections_df = excel_data.parse(sheet_name=2)
        else:
            corrections_df = pd.DataFrame(columns=[
                "Kode Aset", "Tanggal Koreksi", "Jumlah"
            ])

        required_assets = {
            "Kode Aset",
            "Harga Perolehan Awal (Rp)",
            "Tanggal Perolehan",
            "Masa Manfaat (tahun)"
        }

        if not required_assets.issubset(assets_df.columns):
            st.error("Kolom di Sheet 1 tidak valid. Wajib: Kode Aset, Harga Perolehan Awal (Rp), Tanggal Perolehan, Masa Manfaat (tahun).")
            return

        if not capitalizations_df.empty:
            required_caps = {
                "Kode Aset",
                "Tanggal Kapitalisasi",
                "Jumlah",
                "Tambahan Usia"
            }
            if not required_caps.issubset(capitalizations_df.columns):
                st.error("Kolom di Sheet 2 tidak valid. Wajib: Kode Aset, Tanggal Kapitalisasi, Jumlah, Tambahan Usia.")
                return
        else:
            capitalizations_df = pd.DataFrame(columns=[
                "Kode Aset", "Tanggal Kapitalisasi", "Jumlah", "Tambahan Usia"
            ])

        if not corrections_df.empty:
            required_corrs = {
                "Kode Aset",
                "Tanggal Koreksi",
                "Jumlah"
            }
            if not required_corrs.issubset(corrections_df.columns):
                st.error("Kolom di Sheet 3 tidak valid. Wajib: Kode Aset, Tanggal Koreksi, Jumlah.")
                return
        else:
            corrections_df = pd.DataFrame(columns=[
                "Kode Aset", "Tanggal Koreksi", "Jumlah"
            ])

        assets_df["Kode Aset"] = assets_df["Kode Aset"].apply(normalize_kode_aset)
        capitalizations_df["Kode Aset"] = capitalizations_df["Kode Aset"].apply(normalize_kode_aset)
        corrections_df["Kode Aset"] = corrections_df["Kode Aset"].apply(normalize_kode_aset)

        assets_df["Harga Perolehan Awal (Rp)"] = pd.to_numeric(
            assets_df["Harga Perolehan Awal (Rp)"], errors="coerce"
        )
        assets_df["Masa Manfaat (tahun)"] = pd.to_numeric(
            assets_df["Masa Manfaat (tahun)"], errors="coerce"
        )

        if not capitalizations_df.empty:
            capitalizations_df["Jumlah"] = pd.to_numeric(
                capitalizations_df["Jumlah"], errors="coerce"
            )
            capitalizations_df["Tambahan Usia"] = pd.to_numeric(
                capitalizations_df["Tambahan Usia"], errors="coerce"
            )

        if not corrections_df.empty:
            corrections_df["Jumlah"] = pd.to_numeric(
                corrections_df["Jumlah"], errors="coerce"
            )

        assets_df["Tanggal Perolehan"] = assets_df["Tanggal Perolehan"].apply(parse_mixed_excel_date)

        if not capitalizations_df.empty:
            capitalizations_df["Tanggal Kapitalisasi"] = capitalizations_df["Tanggal Kapitalisasi"].apply(parse_mixed_excel_date)

        if not corrections_df.empty:
            corrections_df["Tanggal Koreksi"] = corrections_df["Tanggal Koreksi"].apply(parse_mixed_excel_date)

        aset_valid = assets_df.dropna(subset=["Kode Aset"])
        duplicated_codes = aset_valid[aset_valid["Kode Aset"].duplicated()]["Kode Aset"].unique().tolist()

        if duplicated_codes:
            st.error("Terdapat duplikat Kode Aset pada Sheet 1: " + ", ".join(map(str, duplicated_codes)))
            return

        skipped_rows = []
        anomaly_rows = []
        results = []
        schedules_dict = {}
        total_rows = len(assets_df)

        for idx, asset in assets_df.iterrows():
            alasan = []

            if pd.isna(asset["Kode Aset"]):
                alasan.append("Kode Aset kosong/tidak valid")
            if pd.isna(asset["Harga Perolehan Awal (Rp)"]):
                alasan.append("Harga Perolehan kosong/tidak valid")
            if pd.isna(asset["Tanggal Perolehan"]):
                alasan.append("Tanggal Perolehan kosong/tidak valid")
            if pd.isna(asset["Masa Manfaat (tahun)"]):
                alasan.append("Masa Manfaat kosong/tidak valid")

            if alasan:
                skipped_rows.append({
                    "Baris Excel": idx + 2,
                    "Kode Aset": asset.get("Kode Aset", ""),
                    "Alasan": "; ".join(alasan)
                })
                continue

            asset_code = str(asset["Kode Aset"]).strip()
            initial_cost = float(asset["Harga Perolehan Awal (Rp)"])
            acquisition_date = asset["Tanggal Perolehan"]
            useful_life = float(asset["Masa Manfaat (tahun)"])

            if initial_cost < 0:
                skipped_rows.append({
                    "Baris Excel": idx + 2,
                    "Kode Aset": asset_code,
                    "Alasan": "Harga Perolehan negatif"
                })
                continue

            if useful_life <= 0:
                skipped_rows.append({
                    "Baris Excel": idx + 2,
                    "Kode Aset": asset_code,
                    "Alasan": "Masa Manfaat harus lebih dari 0"
                })
                continue

            if acquisition_date > REPORTING_DATE:
                skipped_rows.append({
                    "Baris Excel": idx + 2,
                    "Kode Aset": asset_code,
                    "Alasan": "Tanggal Perolehan setelah 31/12/2025"
                })
                continue

            asset_caps = []
            if not capitalizations_df.empty:
                asset_cap_rows = capitalizations_df[
                    capitalizations_df["Kode Aset"] == asset_code
                ].to_dict("records")

                for cap in asset_cap_rows:
                    cap_date = parse_mixed_excel_date(cap.get("Tanggal Kapitalisasi"))
                    if pd.isna(cap_date):
                        continue

                    if cap_date < acquisition_date:
                        anomaly_rows.append({
                            "Kode Aset": asset_code,
                            "Jenis Anomali": "Kapitalisasi sebelum induk",
                            "Tanggal Aset": acquisition_date.strftime("%d/%m/%Y"),
                            "Tanggal Transaksi": cap_date.strftime("%d/%m/%Y"),
                            "Keterangan": f"Tanggal kapitalisasi lebih awal dari tanggal perolehan aset induk. Nilai: {cap.get('Jumlah', 0)}"
                        })
                    else:
                        asset_caps.append(cap)

            asset_corrs = []
            if not corrections_df.empty:
                asset_corr_rows = corrections_df[
                    corrections_df["Kode Aset"] == asset_code
                ].to_dict("records")

                for corr in asset_corr_rows:
                    corr_date = parse_mixed_excel_date(corr.get("Tanggal Koreksi"))
                    if pd.isna(corr_date):
                        continue

                    if corr_date < acquisition_date:
                        anomaly_rows.append({
                            "Kode Aset": asset_code,
                            "Jenis Anomali": "Koreksi sebelum induk",
                            "Tanggal Aset": acquisition_date.strftime("%d/%m/%Y"),
                            "Tanggal Transaksi": corr_date.strftime("%d/%m/%Y"),
                            "Keterangan": f"Tanggal koreksi lebih awal dari tanggal perolehan aset induk. Nilai: {corr.get('Jumlah', 0)}"
                        })
                    else:
                        asset_corrs.append(corr)

            schedule = calculate_depreciation_monthly(
                initial_cost=initial_cost,
                acquisition_date=acquisition_date,
                useful_life_years=useful_life,
                reporting_date=REPORTING_DATE,
                capitalizations=asset_caps,
                corrections=asset_corrs
            )

            if schedule:
                last_row = schedule[-1]
                results.append({
                    "Kode Aset": asset_code,
                    "Tanggal Pelaporan": REPORTING_DATE.strftime("%d/%m/%Y"),
                    "Periode Pelaporan": last_row["Periode"],
                    "Penyusutan Bulan Berjalan": last_row["Penyusutan Bulan Berjalan"],
                    "Akumulasi Penyusutan": last_row["Akumulasi Penyusutan"],
                    "Nilai Buku Akhir": last_row["Nilai Buku Akhir"],
                    "Sisa Masa Manfaat (Bulan)": last_row["Sisa Masa Manfaat (Bulan)"],
                })
                schedules_dict[asset_code] = schedule

        results_df = pd.DataFrame(results)
        skipped_df = pd.DataFrame(skipped_rows)
        anomaly_df = pd.DataFrame(anomaly_rows)

        st.markdown("### Ringkasan Proses")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            status_card("Total Baris", total_rows, "status-blue")
        with c2:
            status_card("Berhasil Diproses", len(results), "status-green")
        with c3:
            status_card("Dilewati", len(skipped_rows), "status-yellow")
        with c4:
            status_card("Anomali", len(anomaly_rows), "status-red")

        search_col1, search_col2 = st.columns([2, 1])
        with search_col1:
            search_kode = st.text_input("🔎 Cari Kode Aset", placeholder="Contoh: 211506")
            st.markdown('<div class="search-box-note">Pencarian dipakai untuk hasil, anomali, dan detail per aset.</div>', unsafe_allow_html=True)

        with search_col2:
            filter_anomali = st.selectbox(
                "Filter Anomali",
                options=["Semua", "Hanya Anomali", "Tanpa Anomali"]
            )

        filtered_results_df = results_df.copy()
        filtered_anomaly_df = anomaly_df.copy()

        if search_kode:
            filtered_results_df = filtered_results_df[
                filtered_results_df["Kode Aset"].astype(str).str.contains(search_kode, case=False, na=False)
            ]
            filtered_anomaly_df = filtered_anomaly_df[
                filtered_anomaly_df["Kode Aset"].astype(str).str.contains(search_kode, case=False, na=False)
            ]

        anomaly_asset_codes = set(filtered_anomaly_df["Kode Aset"].astype(str).tolist()) if not filtered_anomaly_df.empty else set()

        if filter_anomali == "Hanya Anomali" and not filtered_results_df.empty:
            filtered_results_df = filtered_results_df[
                filtered_results_df["Kode Aset"].astype(str).isin(anomaly_asset_codes)
            ]

        if filter_anomali == "Tanpa Anomali" and not filtered_results_df.empty:
            filtered_results_df = filtered_results_df[
                ~filtered_results_df["Kode Aset"].astype(str).isin(anomaly_asset_codes)
            ]

        tab1, tab2, tab3 = st.tabs(["📊 Hasil Perhitungan", "📝 Reviu Hasil", "📂 Detail per Aset"])

        with tab1:
            st.markdown("#### Ringkasan Hasil")
            if not filtered_results_df.empty:
                st.dataframe(
                    filtered_results_df.style.format({
                        "Penyusutan Bulan Berjalan": "{:,.2f}",
                        "Akumulasi Penyusutan": "{:,.2f}",
                        "Nilai Buku Akhir": "{:,.2f}",
                        "Sisa Masa Manfaat (Bulan)": "{:,.0f}",
                    }),
                    use_container_width=True
                )
            else:
                st.warning("Tidak ada data hasil yang sesuai filter.")

        with tab2:
            left, right = st.columns(2)

            with left:
                st.markdown("#### Baris yang Dilewati")
                if not skipped_df.empty:
                    if search_kode:
                        skipped_df_show = skipped_df[
                            skipped_df["Kode Aset"].astype(str).str.contains(search_kode, case=False, na=False)
                        ]
                    else:
                        skipped_df_show = skipped_df

                    if not skipped_df_show.empty:
                        st.dataframe(skipped_df_show, use_container_width=True)
                    else:
                        st.info("Tidak ada baris dilewati yang sesuai filter.")
                else:
                    st.success("Tidak ada baris yang dilewati.")

            with right:
                st.markdown("#### Input Aset Tidak Logis / Anomali")
                if not filtered_anomaly_df.empty:
                    st.dataframe(filtered_anomaly_df, use_container_width=True)
                else:
                    st.success("Tidak ada input anomali yang sesuai filter.")

        with tab3:
            st.markdown("#### Detail Jadwal Penyusutan per Aset")

            detail_options = list(schedules_dict.keys())
            if search_kode:
                detail_options = [x for x in detail_options if search_kode.lower() in str(x).lower()]

            if filter_anomali == "Hanya Anomali":
                detail_options = [x for x in detail_options if str(x) in anomaly_asset_codes]

            if filter_anomali == "Tanpa Anomali":
                detail_options = [x for x in detail_options if str(x) not in anomaly_asset_codes]

            if detail_options:
                selected_asset = st.selectbox(
                    "Pilih Kode Aset",
                    options=detail_options
                )

                if selected_asset:
                    detail_df = pd.DataFrame(schedules_dict[selected_asset])
                    st.dataframe(
                        detail_df.style.format({
                            "Kapitalisasi Bulan Ini": "{:,.2f}",
                            "Koreksi Bulan Ini": "{:,.2f}",
                            "Penyusutan Bulan Berjalan": "{:,.2f}",
                            "Akumulasi Penyusutan": "{:,.2f}",
                            "Nilai Buku Akhir": "{:,.2f}",
                            "Sisa Masa Manfaat (Bulan)": "{:,.0f}",
                            "Sisa Masa Manfaat (Tahun)": "{:,.2f}",
                        }),
                        use_container_width=True
                    )
            else:
                st.info("Tidak ada Kode Aset yang sesuai filter.")

        if not results_df.empty or not skipped_df.empty or not anomaly_df.empty:
            st.markdown("---")
            excel_buffer = convert_df_to_excel_with_sheets(
                results,
                schedules_dict,
                skipped_rows=skipped_rows,
                anomaly_rows=anomaly_rows,
                total_rows=total_rows
            )

            st.download_button(
                "📥 Download Hasil Excel",
                excel_buffer,
                "hasil_penyusutan_bulanan_2025.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")


if __name__ == "__main__":
    app()

import streamlit as st
import pandas as pd
from io import BytesIO

REPORTING_DATE = pd.Timestamp("2025-12-31")


def parse_mixed_excel_date(value):
    """
    Menangani:
    - datetime/Timestamp
    - teks tanggal biasa: 14-08-2017, 14/08/2017, 2017-08-14
    - serial date Excel: 42735
    - string dengan spasi tersembunyi
    """
    if pd.isna(value):
        return pd.NaT

    if isinstance(value, pd.Timestamp):
        return value

    text = str(value).strip().replace("\xa0", "").replace("  ", " ")

    if text == "" or text.lower() in ["nan", "none", "nat"]:
        return pd.NaT

    try:
        num = float(text)
        if num > 1000:
            return pd.Timestamp("1899-12-30") + pd.to_timedelta(num, unit="D")
    except Exception:
        pass

    return pd.to_datetime(text, errors="coerce", dayfirst=True)


def normalize_kode_aset(value):
    """
    Menyamakan:
    211506
    211506.0
    ' 211506 '
    menjadi '211506'
    """
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
    """
    Logika:
    - Penyusutan mulai pada bulan perolehan
    - Kapitalisasi diproses pada bulan kapitalisasi
    - Koreksi diproses pada bulan koreksi
    - Tambahan usia diinput dalam TAHUN -> dikali 12 jadi BULAN
    - Sisa masa manfaat maksimum = masa manfaat awal/induk
    """
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


def convert_df_to_excel_with_sheets(results, schedules, skipped_rows=None, total_rows=0):
    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        workbook = writer.book
        money_fmt = workbook.add_format({"num_format": "#,##0.00"})
        int_fmt = workbook.add_format({"num_format": "0"})
        bold_fmt = workbook.add_format({"bold": True})

        # Sheet Ringkasan
        results_df = pd.DataFrame(results)
        results_df.to_excel(writer, index=False, sheet_name="Ringkasan")

        ws_ringkasan = writer.sheets["Ringkasan"]
        ws_ringkasan.set_column("A:A", 20)
        ws_ringkasan.set_column("B:C", 18)
        ws_ringkasan.set_column("D:F", 22, money_fmt)
        ws_ringkasan.set_column("G:G", 20, int_fmt)

        # Sheet detail per aset
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

        # Sheet Reviu Hasil
        ws_reviu = workbook.add_worksheet("Reviu Hasil")
        writer.sheets["Reviu Hasil"] = ws_reviu

        processed_rows = len(results)
        skipped_count = len(skipped_rows) if skipped_rows else 0

        ws_reviu.write(0, 0, "Ringkasan Reviu", bold_fmt)
        ws_reviu.write(2, 0, "Jumlah total baris", bold_fmt)
        ws_reviu.write(2, 1, total_rows, int_fmt)

        ws_reviu.write(3, 0, "Jumlah baris berhasil diproses", bold_fmt)
        ws_reviu.write(3, 1, processed_rows, int_fmt)

        ws_reviu.write(4, 0, "Jumlah baris dilewati", bold_fmt)
        ws_reviu.write(4, 1, skipped_count, int_fmt)

        start_row = 7
        ws_reviu.write(start_row, 0, "Daftar Baris yang Dilewati", bold_fmt)

        skipped_df = pd.DataFrame(skipped_rows if skipped_rows else [], columns=["Baris Excel", "Kode Aset", "Alasan"])
        skipped_df.to_excel(writer, index=False, sheet_name="Reviu Hasil", startrow=start_row + 1)

        ws_reviu.set_column("A:A", 15)
        ws_reviu.set_column("B:B", 20)
        ws_reviu.set_column("C:C", 60)

    buffer.seek(0)
    return buffer.getvalue()


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


def app():
    st.title("📉 Depresiasi GL Bulanan")

    with st.expander("📖 Informasi Batch Tahunan ▼", expanded=False):
        st.markdown("""
        ### Fungsi Batch Bulanan
        1. Unduh template Excel.
        2. Isi data aset, kapitalisasi, dan koreksi.
        3. Unggah file Excel.

        **Tanggal pelaporan otomatis: 31 Desember 2025**

        **Format Excel**

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

        **Catatan**
        - Sheet Kapitalisasi dan Koreksi boleh kosong.
        - Tambahan Usia diisi dalam tahun. Contoh 4 = 48 bulan.
        - Tambahan usia maksimal sampai masa manfaat awal/induk.
        """)

    st.subheader("📥 Download Template Excel")
    template_file = create_template_excel()
    st.download_button(
        "⬇️ Download Template Excel",
        template_file,
        "template_penyusutan_bulanan_2025.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    uploaded_file = st.file_uploader("📤 Unggah File Excel", type=["xlsx"])

    if uploaded_file is not None:
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

                if not capitalizations_df.empty:
                    asset_caps = capitalizations_df[
                        capitalizations_df["Kode Aset"] == asset_code
                    ].to_dict("records")
                else:
                    asset_caps = []

                if not corrections_df.empty:
                    asset_corrs = corrections_df[
                        corrections_df["Kode Aset"] == asset_code
                    ].to_dict("records")
                else:
                    asset_corrs = []

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

            st.info(
                f"Total baris aset: {total_rows} | "
                f"Berhasil diproses: {len(results)} | "
                f"Dilewati: {len(skipped_rows)}"
            )

            results_df = pd.DataFrame(results)

            if not results_df.empty:
                st.dataframe(
                    results_df.style.format({
                        "Penyusutan Bulan Berjalan": "{:,.2f}",
                        "Akumulasi Penyusutan": "{:,.2f}",
                        "Nilai Buku Akhir": "{:,.2f}",
                        "Sisa Masa Manfaat (Bulan)": "{:,.0f}",
                    }),
                    use_container_width=True
                )

                if skipped_rows:
                    st.warning(f"Ada {len(skipped_rows)} baris yang dilewati.")
                    st.dataframe(pd.DataFrame(skipped_rows), use_container_width=True)

                excel_buffer = convert_df_to_excel_with_sheets(
                    results,
                    schedules_dict,
                    skipped_rows=skipped_rows,
                    total_rows=total_rows
                )

                st.download_button(
                    "📥 Download Hasil",
                    excel_buffer,
                    "hasil_penyusutan_bulanan_2025.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Tidak ada data valid yang dapat diproses.")

        except Exception as e:
            st.error(f"❌ Error: {str(e)}")


if __name__ == "__main__":
    app()

import streamlit as st
import pandas as pd
from io import BytesIO

# Tanggal pelaporan otomatis
REPORTING_DATE = pd.Timestamp("2025-12-31")


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
    - Penyusutan dimulai pada bulan perolehan
    - Kapitalisasi diproses pada bulan kapitalisasi
    - Koreksi diproses pada bulan koreksi
    - Tambahan usia diinput dalam TAHUN, lalu dikonversi menjadi BULAN
    - Sisa masa manfaat maksimum = masa manfaat awal/induk
    """
    if capitalizations is None:
        capitalizations = []
    if corrections is None:
        corrections = []

    acquisition_date = pd.to_datetime(acquisition_date, errors="coerce", dayfirst=True)
    reporting_date = pd.to_datetime(reporting_date, errors="coerce")

    if pd.isna(acquisition_date) or pd.isna(reporting_date):
        return []

    if acquisition_date > reporting_date:
        return []

    # Masa manfaat induk dalam bulan
    original_life_months = int(float(useful_life_years) * 12)
    remaining_life_months = original_life_months

    # Nilai buku awal
    book_value = float(initial_cost)
    accumulated_dep = 0.0

    # Kelompokkan kapitalisasi per (tahun, bulan)
    cap_dict = {}
    for cap in capitalizations:
        cap_date = pd.to_datetime(cap.get("Tanggal Kapitalisasi"), errors="coerce", dayfirst=True)
        if pd.notna(cap_date) and cap_date <= reporting_date:
            key = (cap_date.year, cap_date.month)
            cap_dict.setdefault(key, []).append(cap)

    # Kelompokkan koreksi per (tahun, bulan)
    corr_dict = {}
    for corr in corrections:
        corr_date = pd.to_datetime(corr.get("Tanggal Koreksi"), errors="coerce", dayfirst=True)
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

        # 1. Proses kapitalisasi bulan berjalan
        if current_key in cap_dict:
            for cap in cap_dict[current_key]:
                cap_amount = float(cap.get("Jumlah", 0) or 0)

                # Tambahan usia diinput TAHUN -> konversi ke BULAN
                additional_life_years = float(cap.get("Tambahan Usia", 0) or 0)
                additional_life_months = int(additional_life_years * 12)

                kapitalisasi_bulan_ini += cap_amount
                tambahan_usia_bulan_ini += additional_life_months

            # Tambah nilai buku
            book_value += kapitalisasi_bulan_ini

            # Tambah sisa masa manfaat, tapi MAKSIMAL masa manfaat awal/induk
            remaining_life_months = min(
                remaining_life_months + tambahan_usia_bulan_ini,
                original_life_months
            )

        # 2. Proses koreksi bulan berjalan
        if current_key in corr_dict:
            for corr in corr_dict[current_key]:
                corr_amount = float(corr.get("Jumlah", 0) or 0)
                koreksi_bulan_ini += corr_amount

            book_value = max(book_value - koreksi_bulan_ini, 0)

        # 3. Hitung penyusutan bulan berjalan
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

        # Pindah ke bulan berikutnya
        current_month += 1
        if current_month > 12:
            current_month = 1
            current_year += 1

    return schedule


def convert_df_to_excel_with_sheets(results, schedules):
    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        # Sheet ringkasan
        results_df = pd.DataFrame(results)
        results_df.to_excel(writer, index=False, sheet_name="Ringkasan")

        # Sheet detail per kode aset
        for asset_code, schedule in schedules.items():
            schedule_df = pd.DataFrame(schedule)

            valid_sheet_name = (
                str(asset_code)[:31]
                .replace("/", "_")
                .replace("\\", "_")
                .replace(":", "_")
                .replace("*", "_")
                .replace("?", "_")
                .replace("[", "_")
                .replace("]", "_")
            )

            schedule_df.to_excel(writer, sheet_name=valid_sheet_name, startrow=1, index=False)

            worksheet = writer.sheets[valid_sheet_name]
            worksheet.write(0, 0, "Kode Aset")
            worksheet.write(0, 1, asset_code)

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
        """)

    st.subheader("📥 Download Template Excel")
    if st.button("⬇️ Download Template Excel"):
        st.info("Template mengikuti format kolom yang dijelaskan di atas.")

    uploaded_file = st.file_uploader("📤 Unggah File Excel", type=["xlsx"])

    if uploaded_file is not None:
        try:
            # Penting: baca xlsx pakai openpyxl
            excel_data = pd.ExcelFile(uploaded_file, engine="openpyxl")

            assets_df = excel_data.parse(sheet_name=0)
            capitalizations_df = excel_data.parse(sheet_name=1)
            corrections_df = excel_data.parse(sheet_name=2)

            required_assets = {
                "Kode Aset",
                "Harga Perolehan Awal (Rp)",
                "Tanggal Perolehan",
                "Masa Manfaat (tahun)"
            }
            required_caps = {
                "Kode Aset",
                "Tanggal Kapitalisasi",
                "Jumlah",
                "Tambahan Usia"
            }
            required_corrs = {
                "Kode Aset",
                "Tanggal Koreksi",
                "Jumlah"
            }

            if not required_assets.issubset(assets_df.columns):
                st.error("Kolom di Sheet 1 tidak valid. Wajib: Kode Aset, Harga Perolehan Awal (Rp), Tanggal Perolehan, Masa Manfaat (tahun).")
                return

            if not required_caps.issubset(capitalizations_df.columns):
                st.error("Kolom di Sheet 2 tidak valid. Wajib: Kode Aset, Tanggal Kapitalisasi, Jumlah, Tambahan Usia.")
                return

            if not required_corrs.issubset(corrections_df.columns):
                st.error("Kolom di Sheet 3 tidak valid. Wajib: Kode Aset, Tanggal Koreksi, Jumlah.")
                return

            # Normalisasi teks
            assets_df["Kode Aset"] = assets_df["Kode Aset"].astype(str).str.strip()
            capitalizations_df["Kode Aset"] = capitalizations_df["Kode Aset"].astype(str).str.strip()
            corrections_df["Kode Aset"] = corrections_df["Kode Aset"].astype(str).str.strip()

            # Konversi numerik
            assets_df["Harga Perolehan Awal (Rp)"] = pd.to_numeric(
                assets_df["Harga Perolehan Awal (Rp)"], errors="coerce"
            )
            assets_df["Masa Manfaat (tahun)"] = pd.to_numeric(
                assets_df["Masa Manfaat (tahun)"], errors="coerce"
            )

            capitalizations_df["Jumlah"] = pd.to_numeric(
                capitalizations_df["Jumlah"], errors="coerce"
            )
            capitalizations_df["Tambahan Usia"] = pd.to_numeric(
                capitalizations_df["Tambahan Usia"], errors="coerce"
            )

            corrections_df["Jumlah"] = pd.to_numeric(
                corrections_df["Jumlah"], errors="coerce"
            )

            # Konversi tanggal
            assets_df["Tanggal Perolehan"] = pd.to_datetime(
                assets_df["Tanggal Perolehan"], errors="coerce", dayfirst=True
            )
            capitalizations_df["Tanggal Kapitalisasi"] = pd.to_datetime(
                capitalizations_df["Tanggal Kapitalisasi"], errors="coerce", dayfirst=True
            )
            corrections_df["Tanggal Koreksi"] = pd.to_datetime(
                corrections_df["Tanggal Koreksi"], errors="coerce", dayfirst=True
            )

            results = []
            schedules = {}

            for _, asset in assets_df.iterrows():
                if (
                    pd.isna(asset["Kode Aset"]) or
                    pd.isna(asset["Harga Perolehan Awal (Rp)"]) or
                    pd.isna(asset["Tanggal Perolehan"]) or
                    pd.isna(asset["Masa Manfaat (tahun)"])
                ):
                    continue

                asset_code = str(asset["Kode Aset"]).strip()
                initial_cost = float(asset["Harga Perolehan Awal (Rp)"])
                acquisition_date = asset["Tanggal Perolehan"]
                useful_life = float(asset["Masa Manfaat (tahun)"])

                if acquisition_date > REPORTING_DATE:
                    st.warning(f"Kode Aset '{asset_code}' dilewati karena tanggal perolehan setelah 31/12/2025.")
                    continue

                asset_caps = capitalizations_df[
                    capitalizations_df["Kode Aset"] == asset_code
                ].to_dict("records")

                asset_corrs = corrections_df[
                    corrections_df["Kode Aset"] == asset_code
                ].to_dict("records")

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
                    schedules[asset_code] = schedule

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

                excel_buffer = convert_df_to_excel_with_sheets(results, schedules)

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

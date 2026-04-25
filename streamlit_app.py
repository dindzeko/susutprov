import streamlit as st
import pandas as pd
from io import BytesIO

REPORTING_DATE = pd.Timestamp("2025-12-31")


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

    acquisition_date = pd.to_datetime(acquisition_date, errors="coerce", dayfirst=True)
    reporting_date = pd.to_datetime(reporting_date)

    if pd.isna(acquisition_date) or acquisition_date > reporting_date:
        return []

    cap_dict = {}
    for cap in capitalizations:
        cap_date = pd.to_datetime(cap.get("Tanggal Kapitalisasi"), errors="coerce", dayfirst=True)
        if pd.notna(cap_date) and cap_date <= reporting_date:
            key = (cap_date.year, cap_date.month)
            cap_dict.setdefault(key, []).append(cap)

    corr_dict = {}
    for corr in corrections:
        corr_date = pd.to_datetime(corr.get("Tanggal Koreksi"), errors="coerce", dayfirst=True)
        if pd.notna(corr_date) and corr_date <= reporting_date:
            key = (corr_date.year, corr_date.month)
            corr_dict.setdefault(key, []).append(corr)

    original_life_months = int(float(useful_life_years) * 12)
    remaining_life_months = original_life_months

    book_value = float(initial_cost)
    accumulated_dep = 0.0

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
                jumlah = float(cap.get("Jumlah", 0) or 0)
                tambahan_usia = float(cap.get("Tambahan Usia", 0) or 0)

                kapitalisasi_bulan_ini += jumlah
                tambahan_usia_bulan_ini += int(tambahan_usia * 12)

            book_value += kapitalisasi_bulan_ini

            remaining_life_months = min(
                remaining_life_months + tambahan_usia_bulan_ini,
                original_life_months
            )

        if current_key in corr_dict:
            for corr in corr_dict[current_key]:
                jumlah = float(corr.get("Jumlah", 0) or 0)
                koreksi_bulan_ini += jumlah

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


def convert_df_to_excel_with_sheets(results, schedules):
    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        results_df = pd.DataFrame(results)
        results_df.to_excel(writer, index=False, sheet_name="Ringkasan")

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
        """)

    st.subheader("📥 Download Template Excel")
    if st.button("⬇️ Download Template Excel"):
        st.markdown("[Download](https://docs.google.com/spreadsheets/d/1b4bueqvZ0vDn7DtKgNK-uVQojLGMM8vQ/edit?usp=drive_link)")

    uploaded_file = st.file_uploader("📤 Unggah File Excel", type=["xlsx"])

    if uploaded_file is not None:
        try:
            excel_data = pd.ExcelFile(uploaded_file)

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
                st.error("Kolom di Sheet 1 tidak valid! Kolom wajib: Kode Aset, Harga Perolehan Awal (Rp), Tanggal Perolehan, Masa Manfaat (tahun).")
                return

            if not required_caps.issubset(capitalizations_df.columns):
                st.error("Kolom di Sheet 2 tidak valid! Kolom wajib: Kode Aset, Tanggal Kapitalisasi, Jumlah, Tambahan Usia.")
                return

            if not required_corrs.issubset(corrections_df.columns):
                st.error("Kolom di Sheet 3 tidak valid! Kolom wajib: Kode Aset, Tanggal Koreksi, Jumlah.")
                return

            assets_df["Kode Aset"] = assets_df["Kode Aset"].astype(str).str.strip()
            capitalizations_df["Kode Aset"] = capitalizations_df["Kode Aset"].astype(str).str.strip()
            corrections_df["Kode Aset"] = corrections_df["Kode Aset"].astype(str).str.strip()

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
                st.dataframe(results_df.style.format({
                    "Penyusutan Bulan Berjalan": "{:,.2f}".format,
                    "Akumulasi Penyusutan": "{:,.2f}".format,
                    "Nilai Buku Akhir": "{:,.2f}".format,
                    "Sisa Masa Manfaat (Bulan)": "{:,.0f}".format,
                }))

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

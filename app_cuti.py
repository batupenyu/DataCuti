# -*- coding: utf-8 -*-
"""Aplikasi rekap cuti tahunan pegawai SMK Negeri 1 Koba (Streamlit)."""

from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

import ambil_cuti_smkn1koba as cuti

BASE_DIR = Path(__file__).parent
DATA_XLSX = BASE_DIR / "data.xlsx"
LIBUR_XLSX = BASE_DIR / "libur.xlsx"

st.set_page_config(
    page_title="Rekap Cuti Tahunan — SMKN 1 Koba",
    page_icon=":material/event_available:",
    layout="wide",
)


@st.cache_data(show_spinner=False)
def proses(data_bytes: bytes, libur_bytes: bytes):
    libur_dates = cuti.load_libur(BytesIO(libur_bytes))
    output = cuti.load_cuti_dataframe(BytesIO(data_bytes), libur_dates)
    rekap = cuti.hitung_rekap(output)
    kartu = cuti.kartu_kendali(output, rekap)
    return output, rekap, kartu


def unduh_excel(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    return buffer.getvalue()


with st.sidebar:
    st.header("Sumber data")
    up_data = st.file_uploader("data.xlsx — rekam cuti", type=["xlsx"])
    up_libur = st.file_uploader("libur.xlsx — hari libur nasional", type=["xlsx"])
    st.caption("Kosongkan kedua kolom unggah untuk memakai file di folder app.")
    st.divider()
    with st.expander("Aturan perhitungan"):
        st.markdown(
            """
            - Yang dihitung hanya jenis **CUTI TAHUNAN**.
            - Sabtu, Minggu, dan hari libur nasional (libur.xlsx) tidak dihitung.
            - Kuota pokok tiap tahun: **12 hari kerja**.
            - Setiap tahun maksimal 6 hari yang di-carry ke tahun berikutnya.
            - Total akumulasi maksimal **12 hari** → kuota maksimal **24 hari**.
            - Kolom *Sisa cuti* = kuota − cuti, jumlah yang masih bisa diambil di tahun berjalan.
            - Kolom *Carryover* = jumlah yang masuk ke kuota tahun berikutnya (maksimal 6).
            """
        )

if up_data is not None:
    data_bytes, sumber_data = up_data.getvalue(), f"unggahan `{up_data.name}`"
elif DATA_XLSX.exists():
    data_bytes, sumber_data = DATA_XLSX.read_bytes(), f"file lokal `{DATA_XLSX.name}`"
else:
    st.error("data.xlsx tidak ditemukan. Unggah melalui panel di samping.")
    st.stop()

if up_libur is not None:
    libur_bytes, sumber_libur = up_libur.getvalue(), f"unggahan `{up_libur.name}`"
elif LIBUR_XLSX.exists():
    libur_bytes, sumber_libur = LIBUR_XLSX.read_bytes(), f"file lokal `{LIBUR_XLSX.name}`"
else:
    st.error("libur.xlsx tidak ditemukan. Unggah melalui panel di samping.")
    st.stop()

with st.spinner("Menghitung rekap cuti..."):
    try:
        output, rekap, kartu = proses(data_bytes, libur_bytes)
    except Exception as exc:
        st.error(f"Gagal memproses data: {exc}")
        st.stop()

st.title(":material/event_available: Rekap Cuti Tahunan")
st.caption(f"Sumber data: {sumber_data} · sumber libur: {sumber_libur}")

if output.empty:
    st.warning("Tidak ada rekam CUTI TAHUNAN pada data.")
    st.stop()

tahun_berjalan = int(rekap["tahun"].max())
rekap_tb = rekap[rekap["tahun"] == tahun_berjalan]

k1, k2, k3, k4 = st.columns(4)
k1.metric("Jumlah pegawai", int(rekap["nama"].nunique()))
k2.metric("Tahun berjalan", str(tahun_berjalan))
k3.metric(f"Cuti diambil {tahun_berjalan}", f"{int(rekap_tb['cuti_current_year'].sum())} hari")
k4.metric(f"Sisa kuota {tahun_berjalan}", f"{int(rekap_tb['sisa_di_tahun_ini'].sum())} hari")

pilihan = st.selectbox("Pegawai", ["Semua"] + sorted(rekap["nama"].unique()))

tab_rekap, tab_detail, tab_kartu = st.tabs(["Rekap per tahun", "Detail cuti", "Kartu Kendali"])

with tab_rekap:
    df = rekap if pilihan == "Semua" else rekap[rekap["nama"] == pilihan]
    tampilan = df.drop(columns=["total_hari"]).rename(columns={
        "nama": "Nama",
        "tahun": "Tahun",
        "cuti_current_year": "Cuti",
        "sisa_tahun_n_minus_1": "Sisa tahun N-1",
        "kuota_current_year": "Kuota tahun ini",
        "sisa_di_tahun_ini": "Sisa cuti",
        "sisa_cuti_current_year": "Carryover",
    })
    st.dataframe(
        tampilan,
        hide_index=True,
        alt="Tabel rekap saldo cuti per pegawai per tahun",
        column_config={
            "No": st.column_config.NumberColumn(format="%d"),
            "Tahun": st.column_config.NumberColumn(format="%d"),
            "Cuti": st.column_config.NumberColumn(format="%d"),
            "Sisa tahun N-1": st.column_config.NumberColumn(format="%d"),
            "Kuota tahun ini": st.column_config.NumberColumn(format="%d"),
            "Sisa cuti": st.column_config.NumberColumn(format="%d"),
            "Carryover": st.column_config.NumberColumn(format="%d"),
        },
    )
    st.download_button(
        "Unduh rekap (.xlsx)",
        unduh_excel(rekap),
        file_name="rekap_cuti_smkn1koba.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        icon=":material/download:",
    )

with tab_detail:
    dt = output if pilihan == "Semua" else output[output["Nama"] == pilihan]
    st.dataframe(
        dt,
        hide_index=True,
        alt="Tabel detail setiap rekam cuti",
        column_config={
            "No": st.column_config.NumberColumn(format="%d"),
            "Tgl_mulai_cuti": st.column_config.DatetimeColumn(format="DD/MM/YYYY"),
            "Tgl_selesai_cuti": st.column_config.DatetimeColumn(format="DD/MM/YYYY"),
            "Tahun_cuti": st.column_config.NumberColumn(format="%d"),
            "Jumlah_hari_cuti": st.column_config.NumberColumn(format="%d"),
        },
    )
    st.download_button(
        "Unduh detail (.xlsx)",
        unduh_excel(output),
        file_name="data_cuti_smkn1koba.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        icon=":material/download:",
    )

with tab_kartu:
    if pilihan == "Semua":
        kk = cuti.kartu_kendali(output, rekap, tahun=tahun_berjalan)
    else:
        kk = cuti.kartu_kendali(output[output["Nama"] == pilihan], rekap[rekap["nama"] == pilihan], tahun=tahun_berjalan)
    st.dataframe(
        kk,
        hide_index=True,
        alt="Kartu kendali cuti per pegawai",
        column_config={
            "No": st.column_config.NumberColumn(format="%d"),
            "Dari Tgl": st.column_config.DatetimeColumn(format="DD/MM/YYYY"),
            "Sampai Tgl": st.column_config.DatetimeColumn(format="DD/MM/YYYY"),
            "ATB": st.column_config.NumberColumn(format="%d"),
            "LHC": st.column_config.NumberColumn(format="%d"),
            "STB": st.column_config.NumberColumn(format="%d"),
        },
    )
    st.download_button(
        "Unduh kartu kendali (.xlsx)",
        unduh_excel(kk),
        file_name="kartu_kendali_cuti.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        icon=":material/download:",
    )
    st.divider()
    st.subheader("Cetak Kartu Kendali (PDF)")
    col1, col2 = st.columns(2)
    with col1:
        pdfNama = st.text_input("Nama", value=pilihan if pilihan != "Semua" else (str(output["Nama"].iloc[0]) if not output.empty else ""))
        pdfNIP = st.text_input("NIP", value="198410152008041002")
        pdfPangkat = st.text_input("Pangkat/Gol. Ruang", value="Pembina, IV/a")
    with col2:
        pdfJabatan = st.text_input("Jabatan", value="Guru Ahli Madya")
        pdfUnit = st.text_input("Unit Kerja", value="SMK Negeri 1 Koba")
        pdfTahun = st.number_input("Tahun", min_value=2000, max_value=2100, value=int(rekap["tahun"].max()), step=1)
    if st.button("Generate PDF"):
        if not pdfNama:
            st.error("Nama wajib diisi.")
        else:
            with st.spinner("Membuat PDF..."):
                try:
                    import tempfile, os as _os
                    tmp = _os.path.join(tempfile.gettempdir(), "kartu_kendali_temp.pdf")
                    cuti.generate_pdf(
                        kk, pdfNama, pdfNIP, pdfPangkat, pdfJabatan, pdfUnit, int(pdfTahun), tmp
                    )
                    with open(tmp, "rb") as f:
                        pdf_bytes = f.read()
                    _os.remove(tmp)
                    st.session_state["pdf_bytes"] = pdf_bytes
                    st.session_state["pdf_file_name"] = f"kartu_kendali_{pdfNama.replace(' ','_')}.pdf"
                    st.success("PDF siap diunduh.")
                except Exception as exc:
                    st.error(f"Gagal membuat PDF: {exc}")
    if "pdf_bytes" in st.session_state:
        st.download_button(
            "Unduh Kartu Kendali PDF",
            st.session_state["pdf_bytes"],
            file_name=st.session_state["pdf_file_name"],
            mime="application/pdf",
            icon=":material/download:",
        )

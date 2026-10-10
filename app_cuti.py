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
PEGAWAI_XLSX = BASE_DIR / "data_pegawai.xlsx"

st.set_page_config(
    page_title="Rekap Cuti Tahunan — SMKN 1 Koba",
    page_icon=":material/event_available:",
    layout="wide",
)


def _file_hash(path: Path) -> str:
    """Hash isi file — dipakai sebagai bagian dari cache key agar cache
    otomatis invalid saat file Excel berubah."""
    import hashlib
    if not path.exists():
        return "missing"
    return hashlib.md5(path.read_bytes()).hexdigest()


@st.cache_data(show_spinner=False)
def proses(data_bytes: bytes, libur_bytes: bytes):
    """Cache di-key berdasarkan byte isi file, jadi otomatis invalid
    ketika pengguna mengunggah file baru atau file lokal berubah."""
    libur_dates = cuti.load_libur(BytesIO(libur_bytes))
    output = cuti.load_cuti_dataframe(BytesIO(data_bytes), libur_dates)
    rekap = cuti.hitung_rekap(output)
    kartu = cuti.kartu_kendali(output, rekap)
    return output, rekap, kartu


@st.cache_data(show_spinner=False)
def muat_pegawai(pegawai_bytes: bytes):
    """Load master pegawai dari byte — cache key = isi file."""
    return cuti.load_pegawai(BytesIO(pegawai_bytes))


def unduh_excel(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("Sumber data")
    up_data = st.file_uploader("data.xlsx — rekam cuti", type=["xlsx"])
    up_libur = st.file_uploader("libur.xlsx — hari libur nasional", type=["xlsx"])
    up_pegawai = st.file_uploader("data_pegawai.xlsx — master pegawai", type=["xlsx"])
    st.caption("Kosongkan kolom unggah untuk memakai file di folder app.")

    if st.button("🔄 Muat Ulang Data (hapus cache)", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.divider()
    with st.expander("Aturan perhitungan"):
        st.markdown(
            """
            - Yang dihitung hanya jenis **CUTI TAHUNAN**.
            - Sabtu, Minggu, dan hari libur nasional (libur.xlsx) tidak dihitung.
            - Kuota pokok tiap tahun: **12 hari kerja**.
            - Setiap tahun maksimal 6 hari yang di-carry ke tahun berikutnya.
            - Total akumulasi maksimal **12 hari** → kuota maksimal **24 hari**.
            - Kolom *Sisa cuti* = kuota − cuti.
            - Kolom *Carryover* = jumlah yang masuk ke kuota tahun berikutnya (maksimal 6).
            """
        )


# ---------------------------------------------------------------------------
# Sumber data
# ---------------------------------------------------------------------------
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

if up_pegawai is not None:
    pegawai_bytes, sumber_pegawai = up_pegawai.getvalue(), f"unggahan `{up_pegawai.name}`"
elif PEGAWAI_XLSX.exists():
    pegawai_bytes, sumber_pegawai = PEGAWAI_XLSX.read_bytes(), f"file lokal `{PEGAWAI_XLSX.name}`"
else:
    st.error("data_pegawai.xlsx tidak ditemukan. Unggah melalui panel di samping.")
    st.stop()


# ---------------------------------------------------------------------------
# Proses
# ---------------------------------------------------------------------------
with st.spinner("Menghitung rekap cuti..."):
    try:
        output, rekap, kartu = proses(data_bytes, libur_bytes)
        pegawai_df = muat_pegawai(pegawai_bytes)
    except Exception as exc:
        st.error(f"Gagal memproses data: {exc}")
        st.stop()

st.title(":material/event_available: Rekap Cuti Tahunan")
st.caption(
    f"Sumber data: {sumber_data} · sumber libur: {sumber_libur} · "
    f"sumber pegawai: {sumber_pegawai}"
)

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

# Daftar pegawai: gabungan master + pemilik cuti, urut sesuai master.
urutan_master = list(pegawai_df["NAMA"]) if not pegawai_df.empty else []
pemilik_cuti = sorted(output["Nama"].dropna().unique())
daftar_pegawai = [n for n in urutan_master if n in pemilik_cuti]
sisa = [n for n in pemilik_cuti if n not in daftar_pegawai]
daftar_pegawai += sisa

pilihan = st.selectbox("Pegawai", ["Semua"] + daftar_pegawai)

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
        kk = cuti.kartu_kendali(output, rekap, nama=pilihan, tahun=tahun_berjalan)

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
    st.caption(
        "Pilih pegawai di atas, lalu klik **Isi Otomatis dari Master Pegawai** untuk "
        "mengambil NIP, pangkat, dan jabatan dari data_pegawai.xlsx."
    )

    # Default nama dari pilihan
    nama_default = pilihan if pilihan != "Semua" else (
        daftar_pegawai[0] if daftar_pegawai else ""
    )

    # Tombol auto-fill dari master
    if st.button("📋 Isi Otomatis dari Master Pegawai", use_container_width=False):
        if nama_default:
            info = cuti.get_pegawai_info(pegawai_df, nama_default)
            st.session_state["pdf_nama"] = info["nama"]
            st.session_state["pdf_nip"] = info["nip"]
            st.session_state["pdf_pangkat"] = info["pangkat"]
            st.session_state["pdf_jabatan"] = info["jabatan"]
            st.session_state["pdf_unit"] = cuti.UNIT_KERJA_DEFAULT
            st.success(f"Form diisi otomatis dari master: {info['nama']}")
        else:
            st.warning("Tidak ada pegawai yang dipilih.")

    # Set nilai default di session state saat pertama kali.
    if "pdf_nama" not in st.session_state:
        info_awal = cuti.get_pegawai_info(pegawai_df, nama_default) if nama_default else {
            "nama": "", "nip": "-", "pangkat": "-", "jabatan": "-"
        }
        st.session_state["pdf_nama"] = info_awal["nama"]
        st.session_state["pdf_nip"] = info_awal["nip"]
        st.session_state["pdf_pangkat"] = info_awal["pangkat"]
        st.session_state["pdf_jabatan"] = info_awal["jabatan"]
        st.session_state["pdf_unit"] = cuti.UNIT_KERJA_DEFAULT

    col1, col2 = st.columns(2)
    with col1:
        pdfNama = st.text_input("Nama", key="pdf_nama")
        pdfNIP = st.text_input("NIP", key="pdf_nip")
        pdfPangkat = st.text_input("Pangkat/Gol. Ruang", key="pdf_pangkat")
    with col2:
        pdfJabatan = st.text_input("Jabatan", key="pdf_jabatan")
        pdfUnit = st.text_input("Unit Kerja", key="pdf_unit")
        pdfTahun = st.number_input(
            "Tahun", min_value=2000, max_value=2100,
            value=int(tahun_berjalan), step=1,
        )

    if st.button("🖨️ Generate PDF"):
        if not pdfNama:
            st.error("Nama wajib diisi.")
        else:
            with st.spinner("Membuat PDF..."):
                try:
                    import tempfile, os as _os
                    tmp = _os.path.join(tempfile.gettempdir(), "kartu_kendali_temp.pdf")

                    # Ambil ulang kartu kendali sesuai nama & tahun di form
                    kk_pdf = cuti.kartu_kendali(
                        output, rekap, nama=pdfNama, tahun=int(pdfTahun)
                    )
                    if kk_pdf.empty:
                        st.warning(
                            f"Tidak ada data cuti untuk **{pdfNama}** pada tahun "
                            f"{int(pdfTahun)}. PDF tetap dibuat dengan tabel kosong."
                        )

                    cuti.generate_pdf(
                        kk_pdf,
                        pdfNama, pdfNIP, pdfPangkat, pdfJabatan, pdfUnit,
                        int(pdfTahun), tmp,
                    )
                    with open(tmp, "rb") as f:
                        pdf_bytes = f.read()
                    _os.remove(tmp)

                    st.session_state["pdf_bytes"] = pdf_bytes
                    st.session_state["pdf_file_name"] = (
                        f"kartu_kendali_{pdfTahun}_"
                        f"{pdfNama.replace(' ', '_')}.pdf"
                    )
                    st.success("PDF siap diunduh.")
                except Exception as exc:
                    st.error(f"Gagal membuat PDF: {exc}")

    if "pdf_bytes" in st.session_state:
        st.download_button(
            "⬇️ Unduh Kartu Kendali PDF",
            st.session_state["pdf_bytes"],
            file_name=st.session_state["pdf_file_name"],
            mime="application/pdf",
            icon=":material/download:",
        )
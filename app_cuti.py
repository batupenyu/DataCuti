# -*- coding: utf-8 -*-
"""Aplikasi rekap cuti tahunan pegawai SMK Negeri 1 Koba (Streamlit)."""

from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

import ambil_cuti_smkn1koba as cuti

BASE_DIR = Path(__file__).parent
DATA_XLSX = BASE_DIR / "data.xlsx"
LIBUR_XLSX = BASE_DIR / "libur.xlsx"
PEGAWAI_XLSX = BASE_DIR / "data_pegawai.xlsx"

# Data kepala sekolah (bisa diubah manual di form).
KEPALA_NAMA = "SYAHRYANTO, S.T.,M.Pd"
KEPALA_PANGKAT = "Pembina Tk.I, IV/b"
KEPALA_NIP = "197708262006041005"

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


@st.cache_data(show_spinner=False)
def muat_pegawai(pegawai_bytes: bytes):
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

urutan_master = list(pegawai_df["NAMA"]) if not pegawai_df.empty else []
pemilik_cuti = sorted(output["Nama"].dropna().unique())
daftar_pegawai = [n for n in urutan_master if n in pemilik_cuti]
sisa = [n for n in pemilik_cuti if n not in daftar_pegawai]
daftar_pegawai += sisa

pilihan = st.selectbox("Pegawai", ["Semua"] + daftar_pegawai)

tab_rekap, tab_detail, tab_kartu, tab_surat = st.tabs(
    ["Rekap per tahun", "Detail cuti", "Kartu Kendali", "Surat Cuti"]
)

# ---------------------------------------------------------------------------
# Tab 1: Rekap
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# Tab 2: Detail
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# Tab 3: Kartu Kendali
# ---------------------------------------------------------------------------
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

    nama_default = pilihan if pilihan != "Semua" else (
        daftar_pegawai[0] if daftar_pegawai else ""
    )

    if st.button("📋 Isi Otomatis dari Master Pegawai", key="kk_autofill"):
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
            value=int(tahun_berjalan), step=1, key="kk_tahun",
        )

    if st.button("🖨️ Generate PDF Kartu Kendali"):
        if not pdfNama:
            st.error("Nama wajib diisi.")
        else:
            with st.spinner("Membuat PDF..."):
                try:
                    import tempfile, os as _os
                    tmp = _os.path.join(tempfile.gettempdir(), "kartu_kendali_temp.pdf")
                    kk_pdf = cuti.kartu_kendali(output, rekap, nama=pdfNama, tahun=int(pdfTahun))
                    if kk_pdf.empty:
                        st.warning(f"Tidak ada data cuti untuk {pdfNama} pada tahun {int(pdfTahun)}.")
                    cuti.generate_pdf(
                        kk_pdf, pdfNama, pdfNIP, pdfPangkat, pdfJabatan, pdfUnit,
                        int(pdfTahun), tmp,
                    )
                    with open(tmp, "rb") as f:
                        pdf_bytes = f.read()
                    _os.remove(tmp)
                    st.session_state["pdf_bytes"] = pdf_bytes
                    st.session_state["pdf_file_name"] = (
                        f"kartu_kendali_{pdfTahun}_{pdfNama.replace(' ', '_')}.pdf"
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

# ---------------------------------------------------------------------------
# Tab 4: Surat Cuti
# ---------------------------------------------------------------------------
with tab_surat:
    st.subheader("Surat Permohonan Cuti Tahunan")
    st.caption(
        "Isi data di bawah untuk menghasilkan PDF surat permohonan cuti. "
        "Nama, NIP, Pangkat/Gol, dan Jabatan dapat diisi otomatis dari master pegawai."
    )

    nama_surat_default = pilihan if pilihan != "Semua" else (
        daftar_pegawai[0] if daftar_pegawai else ""
    )

    # Tombol auto-fill dari master
    if st.button("📋 Isi Otomatis dari Master Pegawai", key="surat_autofill"):
        if nama_surat_default:
            info = cuti.get_pegawai_info(pegawai_df, nama_surat_default)
            st.session_state["surat_nama"] = info["nama"]
            st.session_state["surat_nip"] = info["nip"]
            st.session_state["surat_pangkat_gol"] = cuti.format_pangkat_gol(
                info["pangkat"], info["gol"]
            )
            st.session_state["surat_jabatan"] = info["jabatan"]
            st.session_state["surat_unit"] = cuti.UNIT_KERJA_DEFAULT
            st.success(f"Form diisi otomatis dari master: {info['nama']}")
        else:
            st.warning("Tidak ada pegawai yang dipilih.")

    # Inisialisasi session state dari master saat pertama kali.
    if "surat_nama" not in st.session_state:
        info_awal = cuti.get_pegawai_info(pegawai_df, nama_surat_default) if nama_surat_default else {
            "nama": "", "nip": "-", "pangkat": "-", "gol": "-", "jabatan": "-"
        }
        st.session_state["surat_nama"] = info_awal["nama"]
        st.session_state["surat_nip"] = info_awal["nip"]
        st.session_state["surat_pangkat_gol"] = cuti.format_pangkat_gol(
            info_awal["pangkat"], info_awal["gol"]
        )
        st.session_state["surat_jabatan"] = info_awal["jabatan"]
        st.session_state["surat_unit"] = cuti.UNIT_KERJA_DEFAULT

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Data Pemohon**")
        sNama = st.text_input("Nama", key="surat_nama")
        sNIP = st.text_input("NIP", key="surat_nip")
        sPangkatGol = st.text_input("Pangkat/Gol.Ruang", key="surat_pangkat_gol")
        sJabatan = st.text_input("Jabatan", key="surat_jabatan")
        sUnit = st.text_input("Unit Kerja", key="surat_unit")
    with col2:
        st.markdown("**Detail Cuti**")
        sTanggalSurat = st.date_input(
            "Tanggal Surat", value=date.today(), key="surat_tgl_surat"
        )
        sLamaHari = st.number_input(
            "Lama Cuti (hari kerja)", min_value=1, max_value=365, value=3,
            step=1, key="surat_lama",
        )
        sTglMulai = st.date_input("Tanggal Mulai", value=date.today(), key="surat_mulai")
        sTglSelesai = st.date_input(
            "Tanggal Selesai", value=date.today() + timedelta(days=2), key="surat_selesai"
        )
        sAlasan = st.text_input("Alasan", value="Urusan Keluarga", key="surat_alasan")
        sAlamat = st.text_input(
            "Alamat", value="Komp. Perumahan Guru Berok Koba", key="surat_alamat"
        )
        sKontak = st.text_input(
            "Nomor Contact", value="+62 822-7809-3771", key="surat_kontak"
        )

    with st.expander("Data Kepala Sekolah (untuk blok tanda tangan)"):
        c1, c2 = st.columns(2)
        with c1:
            kNama = st.text_input("Nama Kepala", value=KEPALA_NAMA, key="kepala_nama")
            kPangkat = st.text_input("Pangkat Kepala", value=KEPALA_PANGKAT, key="kepala_pangkat")
        with c2:
            kNIP = st.text_input("NIP Kepala", value=KEPALA_NIP, key="kepala_nip")

    if st.button("🖨️ Generate PDF Surat Cuti"):
        if not sNama:
            st.error("Nama wajib diisi.")
        elif sTglSelesai < sTglMulai:
            st.error("Tanggal selesai tidak boleh sebelum tanggal mulai.")
        else:
            with st.spinner("Membuat surat cuti..."):
                try:
                    import tempfile, os as _os
                    tmp = _os.path.join(tempfile.gettempdir(), "surat_cuti_temp.pdf")
                    cuti.generate_surat_cuti(
                        nama=sNama,
                        nip=sNIP,
                        pangkat_gol=sPangkatGol,
                        jabatan=sJabatan,
                        unit_kerja=sUnit,
                        tanggal_surat=sTanggalSurat,
                        lama_hari=int(sLamaHari),
                        tgl_mulai=sTglMulai,
                        tgl_selesai=sTglSelesai,
                        alasan=sAlasan,
                        alamat=sAlamat,
                        kontak=sKontak,
                        kepala_nama=kNama,
                        kepala_pangkat=kPangkat,
                        kepala_nip=kNIP,
                        path_out=tmp,
                    )
                    with open(tmp, "rb") as f:
                        pdf_bytes = f.read()
                    _os.remove(tmp)
                    st.session_state["surat_pdf_bytes"] = pdf_bytes
                    st.session_state["surat_pdf_file_name"] = (
                        f"surat_cuti_{sNama.replace(' ', '_')}.pdf"
                    )
                    st.success("PDF surat cuti siap diunduh.")
                except Exception as exc:
                    st.error(f"Gagal membuat PDF: {exc}")

    if "surat_pdf_bytes" in st.session_state:
        st.download_button(
            "⬇️ Unduh Surat Cuti PDF",
            st.session_state["surat_pdf_bytes"],
            file_name=st.session_state["surat_pdf_file_name"],
            mime="application/pdf",
            icon=":material/download:",
        )
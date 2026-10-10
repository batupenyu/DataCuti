#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script untuk mengambil data cuti pegawai khusus SMK Negeri 1 Koba
Data source: data.xlsx (Sheet1), libur.xlsx (daftar hari libur nasional),
             data_pegawai.xlsx (master pegawai: NAMA, NIP, GOL, PANGKAT, JABATAN)

Modul ini juga dipakai oleh app_cuti.py (Streamlit); jalankan langsung
untuk mode CLI seperti sebelumnya.
"""

import os
import re
import shutil
from datetime import datetime

import pandas as pd

EXCEL_PATH = os.path.join(os.path.dirname(__file__), "data.xlsx")
LIBUR_PATH = os.path.join(os.path.dirname(__file__), "libur.xlsx")
PEGAWAI_PATH = os.path.join(os.path.dirname(__file__), "data_pegawai.xlsx")
OUTPUT_CSV = os.path.join(os.path.dirname(__file__), "data_cuti_smkn1koba.csv")

JENIS_CUTI = "CUTI TAHUNAN"
UNIT_KERJA_DEFAULT = "SMK Negeri 1 Koba"

QUOTA_PER_YEAR = 12
MAX_CARRYOVER = 6

colNama = 0
colJenis = 1
colTglMulai = 6
colTglSelesai = 7


# ---------------------------------------------------------------------------
# Util
# ---------------------------------------------------------------------------
def parse_date(val):
    if pd.isna(val):
        return None
    if isinstance(val, (pd.Timestamp, datetime)):
        return val
    s = str(val).strip()
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def count_weekdays(start, end, libur_set):
    if pd.isna(start) or pd.isna(end):
        return 0
    days = 0
    current = start
    while current <= end:
        if current.weekday() < 5 and current.date() not in libur_set:
            days += 1
        current += pd.Timedelta(days=1)
    return days


def _norm_nama(s):
    """Normalisasi nama: uppercase, hapus tanda baca, rapikan spasi."""
    if pd.isna(s):
        return ""
    s = str(s).upper()
    s = re.sub(r"[^A-Z0-9\s]", " ", s)
    return " ".join(s.split())


# ---------------------------------------------------------------------------
# Loader
# ---------------------------------------------------------------------------
def load_libur(path_or_buffer):
    """Baca libur.xlsx (header di baris ke-2) menjadi set tanggal libur."""
    libur_df = pd.read_excel(path_or_buffer, header=1)
    libur_df.columns = ["Tanggal", "Keterangan", "Jenis"]
    libur_df = libur_df.dropna(subset=["Tanggal"])
    return set(pd.to_datetime(libur_df["Tanggal"], errors="coerce").dropna().dt.date)


def load_pegawai(path_or_buffer):
    """Baca data_pegawai.xlsx menjadi DataFrame ter-normalisasi.

    Kolom yang diharapkan: NAMA, NIP, GOL, PANGKAT, JABATAN.
    Menambahkan kolom kunci `nama_key` untuk pencocokan.
    """
    df = pd.read_excel(path_or_buffer)
    df.columns = [str(c).strip().upper() for c in df.columns]

    for col in ("NAMA", "NIP", "GOL", "PANGKAT", "JABATAN"):
        if col not in df.columns:
            df[col] = ""

    df["nama_key"] = df["NAMA"].apply(_norm_nama)
    df = df.drop_duplicates(subset=["nama_key"], keep="first").reset_index(drop=True)
    return df


def get_pegawai_info(pegawai_df, nama):
    """Ambil dict info pegawai berdasarkan nama (case-insensitive)."""
    kosong = {"nama": nama, "nip": "-", "gol": "-", "pangkat": "-", "jabatan": "-"}
    if pegawai_df is None or pegawai_df.empty:
        return kosong
    key = _norm_nama(nama)
    row = pegawai_df[pegawai_df["nama_key"] == key]
    if row.empty:
        return kosong
    r = row.iloc[0]
    return {
        "nama": str(r["NAMA"]),
        "nip": str(r["NIP"]) if pd.notna(r["NIP"]) else "-",
        "gol": str(r["GOL"]) if pd.notna(r["GOL"]) else "-",
        "pangkat": str(r["PANGKAT"]) if pd.notna(r["PANGKAT"]) else "-",
        "jabatan": str(r["JABATAN"]) if pd.notna(r["JABATAN"]) else "-",
    }


def cek_pencocokan_nama(output, pegawai_df):
    """Diagnostik: tampilkan nama yang tidak ketemu di data_pegawai.xlsx."""
    nama_cuti = set(output["Nama"].dropna().apply(_norm_nama))
    nama_master = set(pegawai_df["nama_key"])
    tidak_ketemu = sorted(nama_cuti - nama_master)
    if tidak_ketemu:
        print("\n[PERINGATAN] Nama berikut ada di data.xlsx tapi TIDAK ADA di data_pegawai.xlsx:")
        for n in tidak_ketemu:
            print(f"  - {n}")
    else:
        print("\n[OK] Semua nama di data.xlsx cocok dengan data_pegawai.xlsx")
    return tidak_ketemu


# ---------------------------------------------------------------------------
# Data cuti
# ---------------------------------------------------------------------------
def load_cuti_dataframe(path_or_buffer, libur_dates):
    df = pd.read_excel(path_or_buffer, sheet_name="Sheet1", header=None)

    maskJenis = df[colJenis].astype(str).str.upper().str.contains(JENIS_CUTI, na=False)
    filtered = df[maskJenis].copy()

    filtered["tgl_mulai_dt"] = pd.to_datetime(filtered[colTglMulai].apply(parse_date), errors="coerce")
    filtered["tgl_selesai_dt"] = pd.to_datetime(filtered[colTglSelesai].apply(parse_date), errors="coerce")
    filtered["tahun_cuti"] = filtered["tgl_mulai_dt"].dt.year

    filtered["jumlah_hari_cuti"] = filtered.apply(
        lambda row: count_weekdays(row["tgl_mulai_dt"], row["tgl_selesai_dt"], libur_dates), axis=1
    )

    output = pd.DataFrame({
        "No": range(1, len(filtered) + 1),
        "Nama": filtered[colNama],
        "Keterangan": filtered[2],
        "Tgl_mulai_cuti": filtered["tgl_mulai_dt"],
        "Tgl_selesai_cuti": filtered["tgl_selesai_dt"],
        "Tahun_cuti": filtered["tahun_cuti"].astype("Int64"),
        "Jumlah_hari_cuti": filtered["jumlah_hari_cuti"].astype("Int64"),
    })

    output = output.dropna(subset=["Tgl_mulai_cuti", "Tgl_selesai_cuti", "Tahun_cuti"])
    output = output.drop_duplicates(subset=["Nama", "Tgl_mulai_cuti"], keep="first")
    output = output.drop_duplicates(subset=["Nama", "Tgl_mulai_cuti", "Tgl_selesai_cuti"], keep="first")
    output = output.sort_values(by=["Tgl_mulai_cuti", "Nama"], ascending=False).reset_index(drop=True)
    output["No"] = range(1, len(output) + 1)
    return output


def hitung_rekap(output):
    usage_by_year = output.groupby(["Nama", "Tahun_cuti"], as_index=False)["Jumlah_hari_cuti"].sum()
    tahun_berjalan = int(output["Tahun_cuti"].max())

    saldo_rows = []
    for nama, person_usage in usage_by_year.groupby("Nama"):
        usage_map = person_usage.set_index("Tahun_cuti")["Jumlah_hari_cuti"].astype(int).to_dict()
        usage_map = {int(k): int(v) for k, v in usage_map.items()}
        mulai = tahun_berjalan - 2
        akumulasi = 0
        for tahun in range(mulai, tahun_berjalan + 1):
            cuti = usage_map.get(tahun, 0)
            kuota = QUOTA_PER_YEAR + akumulasi
            sisa_pokok = max(QUOTA_PER_YEAR - cuti, 0)
            sisa_akumulasi = max(akumulasi - max(cuti - QUOTA_PER_YEAR, 0), 0)
            carryover_total = sisa_pokok + sisa_akumulasi
            akumulasi_berikutnya = min(carryover_total, MAX_CARRYOVER)
            akumulasi_masuk = akumulasi
            akumulasi = akumulasi_berikutnya
            saldo_rows.append({
                "Nama": nama,
                "tahun": tahun,
                "cuti_current_year": cuti,
                "akumulasi": akumulasi_masuk,
                "kuota_current_year": kuota,
                "sisa_cuti_current_year": akumulasi_berikutnya,
                "sisa_di_tahun_ini": kuota - cuti,
            })

    saldo = pd.DataFrame(saldo_rows)
    rekap = saldo.rename(columns={"Nama": "nama"})
    rekap["total_hari"] = rekap["cuti_current_year"]
    rekap["sisa_tahun_n_minus_1"] = rekap["akumulasi"]
    rekap = rekap[["nama", "tahun", "cuti_current_year", "sisa_tahun_n_minus_1", "kuota_current_year", "sisa_di_tahun_ini", "sisa_cuti_current_year", "total_hari"]]

    for idx, row in rekap.iterrows():
        if row["nama"] in usage_by_year["Nama"].values:
            for n, pm in usage_by_year.groupby("Nama"):
                if n == row["nama"]:
                    um = pm.set_index("Tahun_cuti")["Jumlah_hari_cuti"].astype(int).to_dict()
                    cuti_n1 = um.get(row["tahun"] - 1, 0)
                    cuti_n2 = um.get(row["tahun"] - 2, 0)
                    if cuti_n1 == 0 and cuti_n2 == 0 and row["tahun"] == tahun_berjalan:
                        rekap.at[idx, "sisa_tahun_n_minus_1"] = QUOTA_PER_YEAR
                        rekap.at[idx, "kuota_current_year"] = QUOTA_PER_YEAR * 2
                        rekap.at[idx, "sisa_di_tahun_ini"] = rekap.at[idx, "kuota_current_year"] - row["cuti_current_year"]
                        rekap.at[idx, "sisa_cuti_current_year"] = min(rekap.at[idx, "sisa_di_tahun_ini"], MAX_CARRYOVER)

    rekap = rekap.sort_values(by=["nama", "tahun"]).reset_index(drop=True)
    rekap.insert(0, "No", range(1, len(rekap) + 1))
    return rekap


def kartu_kendali(output, rekap, nama=None, tahun=None):
    """Buat data kartu kendali cuti per pegawai.

    ATB = saldo cuti SEBELUM cuti pada baris tsb diambil (running balance).
    LHC = lama hari cuti pada baris tsb.
    STB = ATB - LHC.
    """
    if tahun is None:
        tahun = int(rekap["tahun"].max())

    rekap_map = rekap.set_index(["nama", "tahun"])[["kuota_current_year"]].to_dict("index")

    rows = []
    source = output
    if nama is not None:
        source = source[source["Nama"] == nama]

    source = source.sort_values(by=["Nama", "Tgl_mulai_cuti"])

    saldo_berjalan = {}
    no_counter = 0
    last_nama = None

    for _, r in source.iterrows():
        if int(r["Tahun_cuti"]) != tahun:
            continue

        peg = r["Nama"]
        if peg != last_nama:
            no_counter = 1
            last_nama = peg
        else:
            no_counter += 1

        if peg not in saldo_berjalan:
            kuota_awal = rekap_map.get((peg, tahun), {}).get("kuota_current_year", 0)
            saldo_berjalan[peg] = int(kuota_awal)

        atb = saldo_berjalan[peg]
        lhc = int(r["Jumlah_hari_cuti"]) if pd.notna(r["Jumlah_hari_cuti"]) else 0
        stb = atb - lhc
        saldo_berjalan[peg] = stb

        rows.append({
            "No": no_counter,
            "Surat Izin Nomor": "",
            "Surat Izin Tanggal": "",
            "Dari Tgl": r["Tgl_mulai_cuti"],
            "Sampai Tgl": r["Tgl_selesai_cuti"],
            "Jenis Cuti": JENIS_CUTI,
            "Paraf": "",
            "ATB": atb,
            "LHC": lhc,
            "STB": stb,
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Output PDF
# ---------------------------------------------------------------------------
def generate_pdf(kartu, nama, nip, pangkat, jabatan, unit_kerja, tahun, path_out):
    """Buat kartu kendali PDF (HTML + weasyprint)."""
    def _fmt_tgl(val):
        if val is None or (hasattr(val, "isna") and val.isna()):
            return "-"
        try:
            if hasattr(val, "strftime"):
                return val.strftime("%d-%m-%Y")
            return str(val)
        except Exception:
            return str(val)

    def _esc(s):
        return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    rows_html = ""
    for _, r in kartu.iterrows():
        rows_html += "<tr>"
        rows_html += f"<td class='num'>{int(r['No'])}</td>"
        rows_html += f"<td>{_esc(r.get('Surat Izin Nomor', '') or '-')}</td>"
        rows_html += f"<td>{_esc(r.get('Surat Izin Tanggal', '') or '-')}</td>"
        rows_html += f"<td class='num'>{_fmt_tgl(r.get('Dari Tgl', ''))}</td>"
        rows_html += f"<td class='num'>{_fmt_tgl(r.get('Sampai Tgl', ''))}</td>"
        rows_html += f"<td>{_esc(r.get('Jenis Cuti', ''))}</td>"
        rows_html += "<td></td>"
        rows_html += f"<td class='num'>{int(r['ATB'])}</td>"
        rows_html += f"<td class='num'>{int(r['LHC'])}</td>"
        rows_html += f"<td class='num'>{int(r['STB'])}</td>"
        rows_html += "</tr>"

    html = f"""<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="utf-8">
<style>
  @page {{ size: A4 landscape; margin: 15mm; }}
  body {{ font-family: Helvetica, Arial, sans-serif; font-size: 8pt; margin: 0; }}
  .title {{ text-align: center; font-size: 11pt; font-weight: bold; margin: 0; }}
  .sub {{ text-align: center; font-size: 9pt; margin: 2px 0; }}
  .info {{ margin: 6px 0; font-size: 8pt; }}
  .info td {{ vertical-align: top; padding: 1px 2px; }}
  .info td:first-child {{ font-weight: bold; width: 35mm; }}
  .info td:nth-child(2) {{ width: 5mm; }}
  table.kk {{ width: 100%; border-collapse: collapse; font-size: 7pt; clear: both; }}
  table.kk th, table.kk td {{
    border: 0.5px solid black;
    padding: 2px 3px;
    text-align: center;
    vertical-align: middle;
  }}
  table.kk th {{ background-color: #e0e0e0; }}
  table.kk td.num {{ text-align: center; }}
  table.kk td.left {{ text-align: left; }}
  .footer-wrap {{ display: flex; justify-content: space-between; align-items: flex-start; margin-top: 8px; font-size: 8pt; padding-right: 3cm; }}
  .legend {{ margin: 0; }}
  .legend td {{ padding: 1px 2px; }}
  .legend td:first-child {{ font-weight: bold; width: 30mm; }}
  .footer {{ text-align: center; line-height: 1.4; margin-left: auto; }}
</style>
</head>
<body>
<div class="title">KARTU KENDALI CUTI PEGAWAI NEGERI SIPIL</div>
<div class="sub">SMK NEGERI 1 KOBA</div>
<div class="sub">CABANG DINAS PENDIDIKAN WILAYAH I</div>
<div class="sub">DINAS PENDIDIKAN PROVINSI KEPULAUAN BANGKA BELITUNG</div>
<div class="sub">TAHUN {tahun}</div>
<table class="info">
<tr><td>Nama</td><td>:</td><td>{_esc(nama)}</td></tr>
<tr><td>NIP</td><td>:</td><td>{_esc(nip)}</td></tr>
<tr><td>Pangkat/Gol. Ruang</td><td>:</td><td>{_esc(pangkat)}</td></tr>
<tr><td>Jabatan</td><td>:</td><td>{_esc(jabatan)}</td></tr>
<tr><td>Unit Kerja</td><td>:</td><td>{_esc(unit_kerja)}</td></tr>
</table>
<table class="kk">
<thead>
<tr>
<th rowspan="2">No.</th>
<th colspan="2">Surat Izin/Surat Keputusan</th>
<th colspan="2">Lamanya</th>
<th rowspan="2">Jenis Cuti</th>
<th rowspan="2">Paraf Pegawai Kepegawaian</th>
<th rowspan="2">ATB</th>
<th rowspan="2">LHC</th>
<th rowspan="2">STB</th>
</tr>
<tr>
<th>Nomor</th>
<th>Tanggal</th>
<th>Dari Tgl</th>
<th>Sampai Tgl</th>
</tr>
</thead>
<tbody>
{rows_html}
</tbody>
</table>
<div class="footer-wrap">
<table class="legend">
<tr><td>ATB</td><td>: Akumulasi Sisa Tahun Berjalan</td></tr>
<tr><td>LHC</td><td>: Lama Hari Cuti</td></tr>
<tr><td>STB</td><td>: Sisa Cuti Tahun Berjalan</td></tr>
<tr><td>Alokasi Awal</td><td></td></tr>
<tr><td>Tahun N</td><td>: 12</td></tr>
<tr><td>Tahun N-1</td><td>: 6</td></tr>
<tr><td>Tahun N-2</td><td>: 6</td></tr>
<tr><td>Total Alokasi Awal</td><td>: 24</td></tr>
</table>
<div class="footer">
Mengetahui,<br>
Kepala SMK Negeri 1 Koba<br>
<div style="height: 2cm;"></div>
Syahryanto, S.T.,M.Pd<br>
Pembina Tk.I, IV/b<br>
NIP. 197708262006041005
</div>
</div>
</body>
</html>
"""

    from weasyprint import HTML
    HTML(string=html).write_pdf(path_out)
    return path_out


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main():
    libur_dates = load_libur(LIBUR_PATH)
    output = load_cuti_dataframe(EXCEL_PATH, libur_dates)
    pegawai_df = load_pegawai(PEGAWAI_PATH)
    cek_pencocokan_nama(output, pegawai_df)

    output_xlsx = os.path.join(os.path.dirname(__file__), "data_cuti_smkn1koba.xlsx")
    output_rekap_xlsx = os.path.join(os.path.dirname(__file__), "rekap_cuti_smkn1koba.xlsx")

    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
        output.to_excel(writer, index=False)
        ws = writer.sheets["Sheet1"]
        last_row = len(output) + 1
        for row in range(2, last_row + 1):
            ws.cell(row=row, column=4).number_format = "DD/MM/YYYY"
            ws.cell(row=row, column=5).number_format = "DD/MM/YYYY"
    print(f"\nData disimpan di: {output_xlsx}")
    print(f"Total data cuti tahunan: {len(output)} record")

    rekap = hitung_rekap(output)
    rekap.to_excel(output_rekap_xlsx, index=False, engine="openpyxl")
    print(f"Rekap disimpan di: {output_rekap_xlsx}")
    print(f"Total rekap: {len(rekap)} record")


def generate_semua_kartu(tahun=None, out_dir=None, bersihkan=False):
    """Generate kartu kendali PDF untuk SEMUA pegawai yang punya cuti di tahun tsb."""
    libur_dates = load_libur(LIBUR_PATH)
    output = load_cuti_dataframe(EXCEL_PATH, libur_dates)
    rekap = hitung_rekap(output)
    pegawai_df = load_pegawai(PEGAWAI_PATH)

    if tahun is None:
        tahun = int(rekap["tahun"].max())
    if out_dir is None:
        out_dir = os.path.join(os.path.dirname(__file__), "kartu_kendali")

    if bersihkan and os.path.isdir(out_dir):
        shutil.rmtree(out_dir)
    os.makedirs(out_dir, exist_ok=True)

    pemilik_cuti = output[output["Tahun_cuti"] == tahun]["Nama"].dropna().unique()
    urutan_master = list(pegawai_df["NAMA"]) if not pegawai_df.empty else []

    def _urutan(nama):
        try:
            return urutan_master.index(nama)
        except ValueError:
            return 10**9

    pemilik_cuti = sorted(pemilik_cuti, key=_urutan)

    hasil = []
    for nama_peg in pemilik_cuti:
        kartu = kartu_kendali(output, rekap, nama=nama_peg, tahun=tahun)
        if kartu.empty:
            continue

        info = get_pegawai_info(pegawai_df, nama_peg)
        safe_name = "".join(c if c.isalnum() else "_" for c in str(info["nama"])).strip("_")
        path_out = os.path.join(out_dir, f"kartu_kendali_{tahun}_{safe_name}.pdf")

        generate_pdf(
            kartu,
            nama=info["nama"],
            nip=info["nip"],
            pangkat=info["pangkat"],
            jabatan=info["jabatan"],
            unit_kerja=UNIT_KERJA_DEFAULT,
            tahun=tahun,
            path_out=path_out,
        )
        hasil.append(path_out)
        print(f"[OK] {info['nama']} (NIP: {info['nip']}) -> {path_out}")

    return hasil


if __name__ == "__main__":
    main()

    print("\n=== Generate kartu kendali semua pegawai ===")
    pdfs = generate_semua_kartu(bersihkan=True)
    print(f"\nTotal PDF dibuat: {len(pdfs)}")
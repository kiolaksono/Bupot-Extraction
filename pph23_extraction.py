import re
import sys
from pathlib import Path
from datetime import datetime

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------
# Konfigurasi
# ----------------------------------------------------------------------

OUTPUT_FILENAME = "Rekap_BuktiPotong.xlsx"

HEADERS = [
    "No", "Masa Pajak", "Tahun Pajak", "Pembetulan", "Sifat",
    "Nomor Bukti Potong", "NPWP Penerima", "Nama Penerima", "NITKU Penerima",
    "Jenis Fasilitas", "Jenis PPh", "Kode Objek Pajak", "Objek Pajak",
    "Jumlah Penghasilan Bruto", "Tarif (%)", "PPh Dipotong", "Kode Pasal",
    "Jenis Dokumen", "Tanggal Dokumen", "Nomor Dokumen", "NPWP Pemotong",
    "NITKU Pemotong", "Nama Pemotong", "Tanggal Bukti Potong",
    "Nama Penandatangan",
]

BULAN_ID = {
    "januari": 1, "februari": 2, "maret": 3, "april": 4, "mei": 5,
    "juni": 6, "juli": 7, "agustus": 8, "september": 9,
    "oktober": 10, "november": 11, "desember": 12,
}


# ----------------------------------------------------------------------
# Util
# ----------------------------------------------------------------------

def to_number(s):
    """'162.939.049' -> 162939049 (int). Kosong -> None."""
    if not s:
        return None
    s = s.strip().replace(".", "").replace(",", ".")
    try:
        val = float(s)
        return int(val) if val.is_integer() else val
    except ValueError:
        return None


def to_date(s):
    """'24 Juli 2026' -> datetime(2026, 7, 24). Kalau gagal -> teks asli."""
    if not s:
        return None
    s = s.strip()
    m = re.match(r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", s)
    if not m:
        return s
    day, bulan_txt, year = m.groups()
    month = BULAN_ID.get(bulan_txt.lower())
    if not month:
        return s
    try:
        return datetime(int(year), month, int(day))
    except ValueError:
        return s


def get_field(pattern, text, group=1, flags=re.MULTILINE):
    m = re.search(pattern, text, flags)
    return m.group(group).strip() if m else None


# ----------------------------------------------------------------------
# Ekstraksi satu file PDF
# ----------------------------------------------------------------------

def extract_bukpot(pdf_path: Path) -> dict:
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""

    if "BUKTI PEMOTONGAN DAN/ATAU PEMUNGUTAN PPH" not in text.upper():
        raise ValueError("Bukan format Bukti Potong PPh Unifikasi (BPPU) yang dikenali.")

    data = {}

    # --- Header: Nomor, Masa Pajak, Sifat, Status/Pembetulan ---
    m = re.search(
        r"^(\S+)\s+(\d{2}-\d{4})\s+(TIDAK\s+FINAL|FINAL)\s+(.+)$",
        text, re.MULTILINE,
    )
    if not m:
        raise ValueError("Baris header NOMOR/MASA PAJAK/SIFAT tidak ditemukan.")
    nomor_bukpot, masa_pajak, sifat, status = m.groups()
    data["Nomor Bukti Potong"] = nomor_bukpot.strip()
    data["Masa Pajak"] = masa_pajak.strip()
    data["Tahun Pajak"] = masa_pajak.strip().split("-")[-1]
    data["Sifat"] = re.sub(r"\s+", " ", sifat.strip())
    data["Pembetulan"] = status.strip()

    # --- A. Penerima Penghasilan ---
    data["NPWP Penerima"] = get_field(r"A\.1\s+NPWP\s*/\s*NIK\s*:\s*(\S+)", text)
    data["Nama Penerima"] = get_field(r"A\.2\s+NAMA\s*:\s*(.+)", text)
    data["NITKU Penerima"] = get_field(
        r"A\.3.*?:\s*(.+)", text, flags=re.MULTILINE | re.DOTALL
    )
    if data["NITKU Penerima"]:
        # potong di baris pertama saja (setelah ":") sebelum section B.
        data["NITKU Penerima"] = data["NITKU Penerima"].split("\n")[0].strip()

    # --- B.1, B.2 ---
    data["Jenis Fasilitas"] = get_field(r"B\.1\s+Jenis Fasilitas\s*:\s*(.+)", text)
    data["Jenis PPh"] = get_field(r"B\.2\s+Jenis PPh\s*:\s*(.+)", text)

    # --- Tabel B.3 - B.7 (Kode Objek Pajak / Objek Pajak / DPP / Tarif / PPh) ---
    m = re.search(
        r"B\.3\s+B\.4\s+B\.5\s+B\.6\s+B\.7\s*\n"
        r"(\d{2}-\d{3}-\d{2,3})\s+"      # kode objek pajak
        r"(.+?)\s+"                       # objek pajak (bisa multi-baris)
        r"([\d.,]+)\s+"                   # DPP
        r"(\d+(?:[.,]\d+)?)\s+"           # tarif %
        r"([\d.,]+)\s*\n"                 # PPh dipotong
        r"B\.8",
        text, re.DOTALL,
    )
    if not m:
        raise ValueError("Tabel Kode Objek Pajak / DPP / Tarif / PPh (B.3-B.7) tidak ditemukan.")
    kode, objek, dpp, tarif, pph = m.groups()
    data["Kode Objek Pajak"] = kode.strip()
    data["Objek Pajak"] = re.sub(r"[ \t]+", " ", objek.strip())
    data["Jumlah Penghasilan Bruto"] = to_number(dpp)
    data["Tarif (%)"] = to_number(tarif)
    data["PPh Dipotong"] = to_number(pph)
    data["Kode Pasal"] = ""  # tidak ada di format BPPU ini

    # --- B.8, B.9 ---
    data["Jenis Dokumen"] = get_field(r"Jenis Dokumen\s*:\s*(.+?)\s+Tanggal\s*:", text)
    tgl_dok = get_field(r"Jenis Dokumen\s*:.+?Tanggal\s*:\s*(.+)", text)
    data["Tanggal Dokumen"] = to_date(tgl_dok)
    data["Nomor Dokumen"] = get_field(r"B\.9\s+Nomor Dokumen\s*:\s*(.+)", text)

    # --- C. Pemotong/Pemungut ---
    data["NPWP Pemotong"] = get_field(r"C\.1\s+NPWP\s*/\s*NIK\s*:\s*(\S+)", text)
    data["NITKU Pemotong"] = get_field(
        r"C\.2.*?:\s*(.+)", text, flags=re.MULTILINE | re.DOTALL
    )
    if data["NITKU Pemotong"]:
        data["NITKU Pemotong"] = data["NITKU Pemotong"].split("\n")[0].strip()
    data["Nama Pemotong"] = get_field(
        r"C\.3.*?:\s*(.+)", text, flags=re.MULTILINE | re.DOTALL
    )
    if data["Nama Pemotong"]:
        data["Nama Pemotong"] = data["Nama Pemotong"].split("\n")[0].strip()
    tgl_bukpot = get_field(r"C\.4\s+TANGGAL\s*:\s*(.+)", text)
    data["Tanggal Bukti Potong"] = to_date(tgl_bukpot)
    data["Nama Penandatangan"] = get_field(r"C\.5\s+NAMA PENANDATANGAN\s*:\s*(.+)", text)

    return data


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------

def main():
    folder = Path(getattr(sys, "_MEIPASS", None) or Path(__file__).resolve().parent)
    pdf_files = sorted(folder.glob("*.pdf"))

    if not pdf_files:
        print("Tidak ada file PDF ditemukan di folder ini.")
        input("Tekan ENTER untuk keluar...")
        return

    print(f"Ditemukan {len(pdf_files)} file PDF. Memproses...\n")

    rows = []
    failed = []

    for i, pdf_path in enumerate(pdf_files, start=1):
        try:
            data = extract_bukpot(pdf_path)
            rows.append(data)
            print(f"  [OK]    {pdf_path.name}")
        except Exception as e:
            failed.append((pdf_path.name, str(e)))
            print(f"  [GAGAL] {pdf_path.name} -> {e}")

    # ------------------------------------------------------------
    # Tulis Excel
    # ------------------------------------------------------------
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="305496")
    ws.append(HEADERS)
    for col_idx in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"

    for no, data in enumerate(rows, start=1):
        row_values = [no] + [data.get(h, "") for h in HEADERS[1:]]
        ws.append(row_values)

    # lebar kolom otomatis (sederhana)
    for col_idx, header in enumerate(HEADERS, start=1):
        letter = get_column_letter(col_idx)
        max_len = max(len(header), 12)
        ws.column_dimensions[letter].width = min(max_len + 4, 40)

    # format tanggal & angka
    date_cols = [HEADERS.index("Tanggal Dokumen") + 1, HEADERS.index("Tanggal Bukti Potong") + 1]
    num_cols = [HEADERS.index("Jumlah Penghasilan Bruto") + 1, HEADERS.index("PPh Dipotong") + 1]
    for r in range(2, ws.max_row + 1):
        for c in date_cols:
            cell = ws.cell(row=r, column=c)
            if isinstance(cell.value, datetime):
                cell.number_format = "DD-MMM-YYYY"
        for c in num_cols:
            cell = ws.cell(row=r, column=c)
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0"

    # sheet daftar gagal (kalau ada)
    if failed:
        ws2 = wb.create_sheet("Gagal")
        ws2.append(["Nama File", "Alasan"])
        for cell in ws2[1]:
            cell.font = header_font
            cell.fill = header_fill
        for name, reason in failed:
            ws2.append([name, reason])
        ws2.column_dimensions["A"].width = 45
        ws2.column_dimensions["B"].width = 60

    out_path = folder / OUTPUT_FILENAME
    wb.save(out_path)

    print(f"\nSelesai. {len(rows)} berhasil, {len(failed)} gagal.")
    print(f"Hasil disimpan di: {out_path}")
    input("\nTekan ENTER untuk keluar...")


if __name__ == "__main__":
    main()

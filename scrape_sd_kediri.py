import time
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

BASE = "https://referensi.data.kemendikdasmen.go.id/pendidikan/dikdas"

# jf/5 = jenjang SD
KAB_KEDIRI = {
    "051301": "Kras", "051302": "Ringinrejo", "051303": "Ngancar",
    "051304": "Kepung", "051305": "Puncu", "051306": "Plosoklaten",
    "051307": "Wates", "051308": "Kandat", "051309": "Ngadiluwih",
    "051310": "Mojo", "051311": "Semen", "051312": "Banyakan",
    "051313": "Tarokan", "051314": "Grogol", "051315": "Gampengrejo",
    "051316": "Gurah", "051317": "Pagu", "051318": "Papar",
    "051319": "Plemahan", "051320": "Purwoasri", "051321": "Kunjang",
    "051322": "Pare", "051323": "Kandangan", "051324": "Kayen Kidul",
    "051325": "Ngasem", "051326": "Badas",
}

KOTA_KEDIRI = {
    "056301": "Mojoroto",
    "056302": "Kota Kediri",
    "056303": "Pesantren",
}

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
}


#ambil daftar SD untuk satu kecamatan (format: nama sekolah + alamat)
def scrape_kecamatan(kode, nama_kecamatan):
    url = f"{BASE}/{kode}/3/jf/5/all"
    resp = requests.get(url, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    rows = []
    table = soup.find("table")
    if not table:
        print(f"  [WARNING] tabel tidak ditemukan untuk {nama_kecamatan}")
        return rows

    for tr in table.find_all("tr")[1:]:  # skip header
        cols = tr.find_all("td")
        if len(cols) < 4:
            continue
        nama_sekolah = cols[2].get_text(strip=True)
        alamat = cols[3].get_text(strip=True)
        if nama_sekolah:
            rows.append((nama_sekolah, alamat, nama_kecamatan))
    return rows


def main():
    all_rows = []

    print("Mengambil data Kabupaten Kediri...")
    for kode, nama in KAB_KEDIRI.items():
        print(f"  - {nama} ({kode})")
        all_rows.extend(scrape_kecamatan(kode, nama))
        time.sleep(0.5)  # jeda sopan ke server

    print("Mengambil data Kota Kediri...")
    for kode, nama in KOTA_KEDIRI.items():
        print(f"  - {nama} ({kode})")
        all_rows.extend(scrape_kecamatan(kode, nama))
        time.sleep(0.5)

    print(f"\nTotal sekolah terkumpul: {len(all_rows)}")

    
    wb = Workbook()
    ws = wb.active
    ws.title = "Data SD"

    headers = ["Nomor", "Nama Sekolah", "Nama Jalan/Gang", "Kecamatan"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    for i, (nama_sekolah, alamat, kecamatan) in enumerate(all_rows, start=1):
        ws.append([i, nama_sekolah, alamat, kecamatan])

    
    widths = [8, 45, 45, 20]
    for col_letter, w in zip("ABCD", widths):
        ws.column_dimensions[col_letter].width = w

    wb.save("sd_kediri.xlsx")
    print("File tersimpan sebagai sd_kediri.xlsx")


if __name__ == "__main__":
    main()

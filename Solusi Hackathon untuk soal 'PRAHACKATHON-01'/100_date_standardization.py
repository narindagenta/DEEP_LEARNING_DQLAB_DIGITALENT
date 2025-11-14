import sys
import pandas as pd
import re
from datetime import datetime
import time

def normalize_tanggal_transaksi(input_xlsx_path: str, output_xlsx_path: str) -> None:
    """
    Membaca file Excel dan menuliskan file baru dengan format tanggal transaksi
    distandarkan menjadi dd-MM-yyyy. Versi debug (ada print log).
    """
    import pandas as pd
    import re
    from datetime import datetime
    df = pd.read_excel(input_xlsx_path)

    bulan_map = {
        'jan': '01', 'januari': '01', 'january': '01',
        'feb': '02', 'februari': '02', 'february': '02',
        'mar': '03', 'maret': '03', 'march': '03',
        'apr': '04', 'april': '04',
        'mei': '05', 'may': '05',
        'jun': '06', 'juni': '06', 'june': '06',
        'jul': '07', 'juli': '07', 'july': '07',
        'agu': '08', 'agus': '08', 'agustus': '08', 'august': '08', 'aug': '08',
        'sep': '09', 'sept': '09', 'september': '09',
        'okt': '10', 'oct': '10', 'oktober': '10', 'october': '10',
        'nov': '11', 'november': '11',
        'des': '12', 'dec': '12', 'desember': '12', 'december': '12'
    }

    def parse_tanggal(val):
        if pd.isna(val):
            return val

        s = str(val).strip().lower()
        print(f"\n🔹 Input asli: {val}")

        # Hilangkan simbol umum (‘ ’ ' ` , . / -) → ubah jadi spasi
        s = re.sub(r"[‘’'`,./\-]", " ", s)
        s = re.sub(r"\s+", " ", s).strip()
        print(f"➡ Bersih karakter aneh: {s}")

        # Format seperti 2024, 22 Mei → ubah ke "22 Mei 2024"
        s = re.sub(r"^(\d{4})\s*,?\s*(\d{1,2})\s+([a-zA-Z]+)$", r"\2 \3 \1", s)

        # Format "Mei 22 2024" → ubah ke "22 Mei 2024"
        s = re.sub(r"^([a-zA-Z]+)\s+(\d{1,2})\s+(\d{4})$", r"\2 \1 \3", s)

        # Format "22 Mei '24" → "22 Mei 2024"
        s = re.sub(r"(\d{1,2})\s+([a-zA-Z]+)\s+['’`]?(\d{2})\b", r"\1 \2 20\3", s)
        print(f"➡ Setelah normalisasi umum: {s}")

        # Ganti bulan teks jadi angka
        tokens = s.split()
        if len(tokens) >= 3:
            d, m, y = tokens[0], tokens[1], tokens[2]
            mkey = m[:3]
            if mkey in bulan_map:
                m = bulan_map[mkey]
                s = f"{d}-{m}-{y}"

        # Jika hanya 2 digit tahun
        s = re.sub(r"(\d{1,2})-(\d{1,2})-(\d{2})$", r"\1-\2-20\3", s)

        # Coba parse dengan pandas
        try:
            dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
            if pd.notna(dt):
                hasil = dt.strftime("%d-%m-%Y")
                print(f"✅ Dikonversi: {hasil}")
                return hasil
        except Exception as e:
            print(f"⚠️ Gagal pandas.to_datetime: {e}")

        # Coba parse format numerik umum (22.08.2024 / 22 08 2024)
        s2 = re.sub(r"[^\d]", " ", s)
        s2 = re.sub(r"\s+", " ", s2).strip()
        parts = s2.split()
        if len(parts) == 3:
            try:
                d, m, y = map(int, parts)
                if y < 100:
                    y += 2000
                dt = datetime(y, m, d)
                hasil = dt.strftime("%d-%m-%Y")
                print(f"✅ Dikonversi (numeric): {hasil}")
                return hasil
            except Exception as e:
                print(f"⚠️ Gagal konversi manual: {e}")

        print(f"❌ Gagal mengonversi, dikembalikan apa adanya.")
        return val

    # Terapkan fungsi hanya pada kolom 'tanggal transaksi'
    for col in df.columns:
        if col.strip().lower() == "tanggal transaksi":
            print(f"\n=== Memproses kolom: {col} ===")
            df[col] = df[col].apply(parse_tanggal)
            break

     # Menulis ke file Excel output dengan sheet 'transaksi' dan tanpa index
    with pd.ExcelWriter(output_xlsx_path) as writer:
        df.to_excel(writer, index=False, sheet_name="transaksi")
        
    print(f"\n✅ Proses selesai. File hasil tersimpan di: {output_xlsx_path}\n")

# ======================
# MAIN: Jalankan Manual
# ======================
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("⚠️  Gunakan format: python date_standardization.py <nama_file.xlsx>")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = input_file  # hasil ditulis ke file yang sama

    print(f"\n=== Normalisasi Tanggal Transaksi ===")
    normalize_tanggal_transaksi(input_file, output_file)


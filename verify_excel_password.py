import msoffcrypto
import time
import io

EXCEL_FILE = "uji_excel.xlsx"
WORDLIST = "wordlist.txt"

def main():
    start = time.time()
    attempts = 0

    with open(WORDLIST, "r", encoding="utf-8") as wl:
        for pw in wl:
            password = pw.strip()
            if not password:
                continue

            attempts += 1
            try:
                with open(EXCEL_FILE, "rb") as f:
                    office = msoffcrypto.OfficeFile(f)
                    office.load_key(password=password)

                    # Validasi password dilakukan dengan mencoba decrypt ke memory (bukan ke file)
                    out = io.BytesIO()
                    office.decrypt(out)

                elapsed = time.time() - start
                print(f"[SUCCESS] Password benar: {password}")
                print(f"Percobaan: {attempts}")
                print(f"Waktu: {elapsed:.3f} detik")
                return

            except Exception:
                print(f"[FAIL] {password}")

    elapsed = time.time() - start
    print("[FAILED] Password tidak ditemukan di wordlist")
    print(f"Total percobaan: {attempts}")
    print(f"Waktu: {elapsed:.3f} detik")

if __name__ == "__main__":
    main()

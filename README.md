# Cek Rating & Update Limit Counterparty

Pilih tanggal → aplikasi mengecek rating tiap counterparty **pada tanggal itu** → file Excel limit Anda
diunduh kembali dengan rating (dan tanggal laporan) yang sudah diganti. Format & rumus file tetap.

## Menjalankan

**Windows**: klik dua kali `jalankan.bat` → browser terbuka otomatis.

Manual: `pip install -r requirements.txt` lalu `streamlit run app.py` (Python 3.10+, Streamlit 1.40+).

## Alur

1. **Tanggal periode** – rating yang diambil = rating terakhir yang terbit pada/sebelum tanggal itu
   (PEFINDO update 9/11/2026, Anda pilih 8/11/2026 → yang diambil rating sebelum 9/11/2026).
2. **File Excel limit** – unggah file terbaru (format seperti `Limit_Countreparty_09102026.xlsx`),
   atau biarkan kosong untuk memakai file terakhir.
3. **Cek rating** – urutan pencarian tiap counterparty:
   1. entitas utama group di **PEFINDO** (otomatis dari pefindo.com)
   2. kalau tidak ada → **Fitch** (tabel Rating Fitch)
   3. kalau dua-duanya tidak ada → **perusahaan anggota group** berikutnya (mis. Salim Group → Indofood → Salim Ivomas)
   4. kalau tetap tidak ada → rating di file tidak diubah
4. **Download Excel** – yang diganti: tanggal laporan, kolom Rating (Non MI), dan Total Aset Investasi bila diisi.
   Limit dihitung ulang otomatis oleh rumus di file saat dibuka di Excel.

## Pengaturan (tab "Sumber Rating", tersimpan di `data/pengaturan.xlsx`)

- **Entitas per group**: daftar perusahaan tiap group + kode PEFINDO-nya, urutan = prioritas.
- **Rating Fitch**: situs Fitch tidak menyediakan akses data terbuka, jadi rating Fitch diisi dengan menempelkan
  judul rilisnya (mis. `12 Mar 2026 Fitch Affirms Bank OCBC NISP at 'AAA(idn)'; Outlook Stable`) – tanggal, rating,
  dan entitas terbaca otomatis. Data awal = rating `(idn)` di file contoh.

## Catatan

- Halaman PEFINDO disimpan di `data/cache_pefindo/` dan dipakai ulang di hari yang sama; bila internet putus,
  data tersimpan terakhir dipakai.
- Rating yang ditarik, atau yang terakhir diperbarui > ±13 bulan sebelum tanggal periode, dianggap tidak berlaku
  sehingga dicari ke sumber berikutnya.
- Opsi "Perbaiki rumus total": di file contoh `TOTAL NON MANAJER INVESTASI = SUM(D7:D47)` tidak menjumlah 3 baris
  terakhir; opsi ini mengubah rentangnya agar mencakup semua baris.

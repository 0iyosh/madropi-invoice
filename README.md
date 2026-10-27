# MADROPI - Sistem Rekapitulasi & Invoice Pengiriman Material Makadam

Aplikasi lokal untuk membuat dan mencetak invoice pengiriman material (makadam) yang **terintegrasi langsung dengan file template Microsoft Excel asli Anda (`template.xlsx`)**.

---

## 🚀 Cara Menjalankan Aplikasi

1. **Cukup klik 2x file [**`Buka-Aplikasi.bat`**](file:///d:/projek/MADROPI/Buka-Aplikasi.bat)**.
2. Server lokal akan aktif dan browser otomatis membuka halaman aplikasi di `http://localhost:3000`.
3. Biarkan jendela hitam konsol tetap terbuka selama Anda menggunakan aplikasi.

---

## 🌟 Keunggulan Integrasi Template Excel Asli

Aplikasi ini menggunakan **Pilihan 1 (Integrasi Template Master Excel)**:

1. **Unduh File Excel Resmi (`.xlsx`):**
   - Saat Anda menekan tombol hijau **"📥 Unduh File Excel (.xlsx)"**, sistem akan mengambil file master `template.xlsx` Anda di latar belakang.
   - Mengisi otomatis lokasi proyek, baris ritase pengiriman, rumus volume, harga tetap Rp 200.000, total, Grand Total, kalimat terbilang, tanggal invoice, dan nama tanda tangan **MADROPI**.
   - **Semua format asli:** jenis huruf (Arial), ukuran font, warna biru pastel header (`#D9EAF7`), warna kuning Grand Total (`#FFFF00`), border tabel, margin cetak A4 potret, dan rumus asli Excel **100% terjaga utuh tanpa meleset sedikit pun**!
   - Hasil unduhan dapat dibuka langsung di Microsoft Excel, lalu disimpan ke PDF (*File > Save As > PDF*) atau dicetak langsung.

2. **Cetak Langsung via Browser:**
   - Tersedia tombol biru **"🖨️ Cetak / Simpan PDF"** yang tampilannya di layar telah diselaraskan persis dengan grid dan layout Excel asli Anda.

---

## 📋 Fitur dan Otomasi Lengkap

1. **Tab "Buat Invoice":**
   - **Pilihan Plat Nomor:** Cukup pilih nomor plat dari dropdown, dimensi Panjang (P), Lebar (L), dan Tinggi (T) langsung terisi otomatis!
   - **Rumus Volume ($m^3$):** $\text{Volume} = \frac{P \times L \times T}{1.000.000}$ (dibulatkan 3 desimal).
   - **Harga:** Tetap Rp 200.000 / $m^3$.
   - **Total:** $\text{Volume} \times 200.000$.
   - **Grand Total & Terbilang:** Terhitung dan tertulis otomatis dalam Bahasa Indonesia.
   - **Catatan Rekening (Tetap):**
     ```text
     Catatan :
     1. Pembayaran dilakukan secara transfer ke rekening tujuan
     No. Rek Bank BCA     723-544-6995 A/N RETNO HANDAYANI
     ```
   - **Tanda Tangan:** `Tangerang, [Tanggal Invoice]` oleh **MADROPI**.

2. **Tab "Master Plat Mobil":**
   - Menambah, mengubah, dan menghapus daftar plat nomor armada beserta dimensi bak bawaan.
   - Sudah terisi awal: `B 9451 GYV`, `BA 9873 BU`, `B 9207 UYT`.

3. **Tab "Riwayat Invoice":**
   - Menyimpan invoice-invoice yang pernah dibuat agar sewaktu-waktu dapat dibuka dan diedit kembali.

4. **Tab "Catatan & Panduan":**
   - Informasi rekening resmi, petunjuk cetak presisi, dan fitur ekspor/impor cadangan data (JSON).

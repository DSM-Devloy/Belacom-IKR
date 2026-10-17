# Migrasi Sistem IKR Belacom ke Vercel - Laporan Konsultasi

## 1. Keputusan yang saya rekomendasikan

| Hal | Rekomendasi |
|---|---|
| Front end | **Pindah ke Vercel sekarang** (statis, nyaris tanpa risiko) |
| Backend | **Jangan di GitHub** (lihat bagian 2). **Fase 1: tetap Apps Script**, dibungkus gateway Vercel. **Fase 2: pindah data ke Postgres** bila WO sudah ribuan atau pengguna bertambah |
| Harga | Sudah dibuatkan: tabel harga + tanggal efektif di Menu Setting (bagian 6) |

## 2. "Backend di GitHub": koreksi penting

GitHub tidak menjalankan backend. GitHub Pages hanya hosting statis; GitHub Actions bukan API untuk aplikasi yang dipakai teknisi real-time. GitHub tetap berguna sebagai **repositori kode** yang tersambung ke Vercel (push = deploy otomatis). Jadi pilihan backend yang nyata hanya tiga.

## 3. Perbandingan opsi backend

Angka kecepatan di bawah adalah **estimasi kasar berdasarkan karakteristik platform, bukan hasil benchmark di data Anda**. Ukur sendiri dengan tab Network browser sebelum dan sesudah migrasi.

| | A. Apps Script + gateway Vercel | B. Vercel Functions + Google Sheets API | C. Vercel Functions + Postgres (Supabase/Neon) |
|---|---|---|---|
| Kecepatan baca list/dashboard | ~1-3 dtk (cold start + baca seluruh sheet) | ~0,4-1 dtk | ~0,05-0,3 dtk |
| Skala data | Melambat linear seiring baris bertambah (`getDataRange` tiap request) | Sama, tapi bisa `batchGet` + cache | Indeks + query, nyaris datar |
| Upload foto | Drive (sudah jalan) | Service account tidak punya kuota My Drive, perlu Shared Drive | Object storage (Supabase/Vercel Blob/R2) |
| PDF laporan | Sudah jalan (HTML ke PDF bawaan) | Perlu Chromium/PDF library di serverless | Sama dengan B |
| Usaha migrasi | **Kecil** (file yang saya buat siap pakai) | Sedang (tulis ulang ~30 fungsi) | Besar (tulis ulang + migrasi data + auth) |
| Risiko | Rendah | Sedang | Sedang-tinggi, tapi hasil akhir terbaik |
| Biaya | Gratis | Gratis | Free tier cukup untuk skala ini |
| Keamanan | Baik setelah gateway + secret | Baik | Terbaik (hash password, RLS) |

**Kesimpulan:** A memberi untung cepat dan aman hari ini, tetapi **tidak** mempercepat backend, karena request tetap berakhir di Apps Script. Percepatan nyata baru datang di Fase 2 (opsi C). Karena front end sudah memanggil satu pintu (`/api/rpc`), Fase 2 bisa dilakukan **fungsi demi fungsi** tanpa menyentuh UI lagi.

## 4. Temuan dari kode Anda saat ini

**Keamanan (prioritas tinggi)**
1. Login hanya dicek di browser (`sessionStorage`). Siapa pun yang bisa memanggil fungsi server bisa menghapus WO atau user tanpa login. Gateway Fase 1 menutup ini: token bertanda tangan, otorisasi per role, dan secret antara Vercel dan Apps Script.
2. Password disimpan **teks biasa** di sheet `Users`, dan `getUsersList` mengirim semuanya ke browser. Perbaikan butuh perubahan UI (hash + reset password), jadwalkan di Fase 2.
3. Default `admin/admin123`, `leader/leader123`, `team/team123` dibuat otomatis. **Ganti sebelum go-live.**
4. Foto dan PDF di Drive diset "siapa pun dengan link". Bila berisi data pelanggan, pertimbangkan akses terbatas.

**Performa (quick win di Apps Script)**
1. `getReportIKRList` dan `exportReportIKRToExcel` memanggil `repRows.find` di dalam loop (O(n x m)). Ganti dengan `Map` berdasarkan No Tiket.
2. Setiap fungsi membaca seluruh sheet. Untuk dashboard, `CacheService` 60 detik sudah memangkas beban.
3. Ada `getReportDetail` ganda di `Code.gs` (yang kedua menimpa yang pertama). Aman, tapi sebaiknya satu.
4. Upload foto dikirim satu per satu dalam base64 (12 foto x beberapa MB). Lihat risiko batas ukuran di bagian 7.

## 5. Isi paket

```
belacom-vercel/
  public/index.html          <- Index.html Anda + fitur harga + shim
  public/js/gas-shim.js      <- meniru google.script.run via fetch (UI lama tak perlu ditulis ulang)
  api/rpc.js                 <- gateway: login, token, role, teruskan ke Apps Script
  vercel.json
  apps-script/Code.gs        <- Code.gs Anda, tarif hardcode diganti tabel harga
  apps-script/Api.gs         <- doPost + daftar fungsi yang boleh dipanggil
```

### Langkah go-live
1. **Apps Script:** ganti `Code.gs`, tambah file `Api.gs`. Project Settings > Script Properties > `API_SECRET` = string acak panjang. Deploy > New deployment > Web app > Execute as *Me* > Access *Anyone*. Salin URL `/exec`.
2. **GitHub:** push folder `belacom-vercel` ke repo.
3. **Vercel:** Import repo. Environment Variables: `GAS_URL` (URL tadi), `GAS_SECRET` (sama dengan `API_SECRET`), `JWT_SECRET` (string acak lain).
4. Tes dengan data nyata di URL preview Vercel. Versi Apps Script lama tetap hidup, jadi bisa dijalankan paralel dan rollback kapan saja.
5. Ganti password default, lalu arahkan pengguna ke domain Vercel.

## 6. Fitur Setting Harga + Tanggal Efektif

**Lokasi:** Menu Setting (khusus Admin), kartu "Pengaturan Harga WO".

**Data:** sheet baru `Pengaturan_Harga` (dibuat otomatis) berisi Tgl Efektif, Tagihan WO, Bayar Team IKR, Revenue Belacom, siapa yang mengubah, kapan, dan catatan. Baris pertama otomatis diisi tarif lama (100.000 / 75.000 / 25.000) sejak awal, sehingga **data historis tidak berubah** saat fitur diaktifkan.

**Aturan hitung:** harga sebuah WO = baris dengan Tgl Efektif terbesar yang tidak melewati **Tgl Visit** WO (jika Tgl Visit kosong, pakai Created At). Contoh yang sudah saya uji:

| Tgl Visit WO | Harga Tagihan (efektif 15 Okt: 110rb, 1 Des: 120rb) |
|---|---|
| 14 Okt 2026 | 100.000 |
| 15 Okt 2026 | 110.000 |
| 20 Nov 2026 | 110.000 |
| 1 Des 2026 | 120.000 |

**Perilaku UI:**
- Status tiap baris: *Berlaku* (sedang aktif), *Terjadwal* (tanggal masa depan), *Riwayat*.
- Saat mengisi tanggal, muncul **pratinjau dampak**: berapa WO yang terpengaruh dan berapa yang sudah Paid.
- Peringatan (tidak memblokir) bila Tagihan != Bayar Team + Revenue.
- Menyimpan di tanggal yang sudah ada = merevisi baris itu. Harga dasar tidak bisa dihapus.
- Dashboard Tagihan, List WO, dan total per team semuanya memakai harga per WO, bukan satu tarif global.
- Bug kecil ikut diperbaiki: di `Index.html` ada fallback `|| 100000`, sehingga harga `0` akan tampil sebagai 100.000. Sudah diganti `?? 0`.

**Keputusan bisnis yang perlu Anda konfirmasi:** saya memakai **Tgl Visit** sebagai tanggal acuan. Alternatifnya Tgl Bayar (harga mengikuti kapan dibayar) atau harga dibekukan saat WO berstatus Done. Bila yang benar bukan Tgl Visit, ubah satu fungsi `priceForWoRow_` di `Code.gs`.

## 7. Risiko dan hal yang perlu diperhatikan

- **Ukuran upload foto:** fungsi serverless Vercel membatasi body request (sekitar 4,5 MB, cek dokumentasi terkini). Foto HP mentah dalam base64 bisa melewatinya. Kompres foto di browser (kanvas, sisi terpanjang ~1280 px, JPEG kualitas 0,8, biasanya 200-400 KB) sebelum `fileToBase64`. Ini **wajib** sebelum go-live.
- **Waktu generate PDF:** bergantung pada Apps Script (dengan 12 foto bisa belasan detik). `vercel.json` mengatur `maxDuration` 60 dtk, tetapi batas maksimum tergantung paket Vercel Anda, silakan cek.
- **Kuota Apps Script:** ada batas eksekusi bersamaan per pengguna dan per hari (cek dokumentasi Google). Untuk tim IKR kecil aman; ini salah satu alasan Fase 2 ke Postgres.
- **Token berlaku 12 jam**, disimpan di `sessionStorage` (seperti sebelumnya). Belum ada pembatasan percobaan login, tambahkan rate limit (mis. Vercel Firewall) bila diakses publik.
- Saya **belum menguji** end-to-end terhadap spreadsheet dan Apps Script Anda (tidak punya akses). Yang sudah diuji: sintaks seluruh file dan logika harga efektif. Lakukan uji di salinan spreadsheet dulu.

## 8. Roadmap Fase 2 (setelah Fase 1 stabil)
1. Cache + perbaikan `Map` di Apps Script (quick win).
2. Pindahkan **foto** ke object storage; simpan hanya URL.
3. Pindahkan tabel `Users` (dengan hash password), lalu `Work_Orders`, `Report_IKR`, `Teams`, `Pengaturan_Harga` ke Postgres, satu fungsi per rilis lewat `api/rpc.js`.
4. Generate PDF di Vercel (Chromium) atau layanan PDF terpisah.
5. Matikan Apps Script bila semua fungsi sudah pindah.

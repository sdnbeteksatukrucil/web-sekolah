# Web & Portal SD Negeri Betek I

Web publik + portal Admin / Guru / Siswa. Frontend statis (GitHub Pages), backend **Supabase** (database, login, penyimpanan foto) — gratis, cepat, tanpa server sendiri.

```
index.html                 Web publik + tombol Masuk (pojok kanan atas)
app.html                   Portal (dashboard & semua menu)
assets/config.js           URL & kunci Supabase (WAJIB diisi)
assets/app.js              Seluruh logika portal
supabase/schema.sql        Tabel + keamanan (RLS) per role
supabase/functions/create-user/index.ts   Fungsi pembuat akun guru/siswa
```

## Cara Pasang (±20 menit)

### 1. Buat project Supabase
1. Daftar di https://supabase.com → **New project** (pilih region Singapore, catat password database).
2. Menu **Project Settings → API**: salin **Project URL** dan **anon public key**.

### 2. Buat database
Menu **SQL Editor → New query**, tempel seluruh isi `supabase/schema.sql`, klik **Run**.

### 3. Matikan pendaftaran bebas
**Authentication → Sign In / Providers → Email**: matikan **Allow new users to sign up** (akun hanya dibuat admin).

### 4. Buat akun Admin pertama
1. **Authentication → Users → Add user → Create new user**: email `admin@betek1.local`, isi password, centang **Auto Confirm User**. (Login memakai username `admin`.)
2. Salin **User UID**-nya, lalu jalankan di SQL Editor:
```sql
insert into profiles(id, role, nama) values ('UID_ADMIN_DI_SINI', 'admin', 'Administrator');
```

### 5. Pasang fungsi pembuat akun
**Edge Functions → Deploy a new function → Via Editor**, nama `create-user`, tempel isi `supabase/functions/create-user/index.ts`, lalu **Deploy**. (Atau lewat CLI: `supabase functions deploy create-user`.)

### 6. Isi konfigurasi
Edit `assets/config.js`: ganti `url` dan `key` dengan nilai dari langkah 1. Kunci *anon* memang boleh publik; keamanan data dijaga aturan RLS di `schema.sql`.

### 7. Publikasikan lewat GitHub
1. Buat repository di GitHub (mis. `web-sekolah`), unggah semua file proyek ini.
2. **Settings → Pages → Branch: main / root → Save**.
3. Setelah ±1 menit web aktif di `https://USERNAME.github.io/web-sekolah/`.
(Domain sendiri: isi *Custom domain* di halaman yang sama.)

### 8. Mulai pakai
1. Buka web → **🔐 Masuk** → username `admin`.
2. **Pengaturan**: isi nama, profil, visi, misi, kontak, logo, foto, berita, pengumuman, galeri.
3. **Data Guru**: tambah guru (isi wali kelas & link Google Drive) → **Akun** untuk membuat username & password.
4. **Data Siswa**: tambah siswa (QR otomatis muncul) → **Akun** (username default = NISN).
5. Guru login, absen manual atau **📷 Scan QR**, input nilai & catatan.

## Fitur tambahan
- **🖨 Cetak Kartu QR** (menu Data Siswa): mencetak kartu QR semua siswa sesuai filter/pencarian (3 kartu per baris, siap potong). Guru hanya mencetak siswa kelasnya.
- **Edit** nilai dan catatan (selain tambah & hapus).
- **Pencarian** nama/NISN dan **halaman per 20 data** pada Data Siswa, Nilai, dan Catatan agar tetap ringan saat data banyak.

## Catatan penting
- Login memakai **username**; sistem menambahkan `@betek1.local` di belakang layar.
- Scan kamera butuh **HTTPS** (GitHub Pages sudah HTTPS).
- Tombol WA membuka WhatsApp dengan pesan terisi; pengiriman tetap dikonfirmasi manual (WA tidak mengizinkan kirim otomatis tanpa API berbayar). Pastikan nomor HP siswa diisi.
- Lupa password guru/siswa: Supabase → Authentication → Users → ubah password.
- Backup: Supabase → Database → Backups, atau ekspor Excel dari tiap menu.
- Batas gratis Supabase cukup untuk 1 sekolah dasar; foto sebaiknya dikompres < 1 MB agar cepat.

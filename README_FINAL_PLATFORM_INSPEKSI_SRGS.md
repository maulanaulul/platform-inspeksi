# Platform Inspeksi SRGS - Final All-in-One v4

Versi ini menggabungkan 3 aplikasi dalam satu project React + Supabase:

1. Sidak Fatigue
2. DRD Driver
3. Inspeksi

## Perbaikan v4

- Landing page awal dengan 3 card aplikasi dihapus.
- Setelah login, user langsung memilih aplikasi/role/site dari halaman sesi kerja.
- Dropdown aplikasi tidak lagi menampilkan aplikasi dobel.
- Nama aplikasi ketiga diganti dari `Inspeksi Unit` menjadi `Inspeksi`.
- Saat memilih `DRD Driver`, sistem membuka aplikasi DRD Driver asli, bukan placeholder.
- Saat memilih `Inspeksi`, sistem membuka aplikasi Inspeksi asli, bukan placeholder.
- Kategori checklist tetap memakai `Inspeksi Unit` dan `Inspeksi Kelayakan Parkiran`, karena itu kategori pemeriksaan, bukan nama aplikasi.

## Cara Pakai

1. Extract ZIP.
2. Pastikan file `.env` sudah dibuat dari `.env.example`.
3. Jalankan patch SQL v4 di Supabase SQL Editor:

```sql
SQL_PATCH_FINAL_V4_CLEAN_APPS_AND_ROUTE.sql
```

4. Jalankan aplikasi:

```powershell
npm install
npm run dev
```

5. Clear browser cache/session lokal:

```javascript
localStorage.clear();
sessionStorage.clear();
location.reload();
```

## Catatan

Jika sebelumnya database sudah punya aplikasi dobel di tabel `applications`, patch SQL v4 akan memindahkan mapping user ke aplikasi canonical lalu menghapus aplikasi duplikat.

## Final v5 UI Polish + Bulk Access Upload
- Halaman launcher 3 kartu tidak dipakai di flow utama; pemilihan aplikasi tetap lewat profile/Ganti Aplikasi.
- Nama aplikasi ketiga diseragamkan menjadi **Inspeksi**.
- Istilah **Lomba** diubah menjadi **Achievement/Pencapaian**.
- Master Driver DRD dapat diedit dan delete/nonaktif.
- Mapping Akses mendukung upload Excel bulk dengan preview sebelum submit.
- UI dipoles ulang agar Sidak Fatigue, DRD Driver, dan Inspeksi lebih senada, modern, dan premium.

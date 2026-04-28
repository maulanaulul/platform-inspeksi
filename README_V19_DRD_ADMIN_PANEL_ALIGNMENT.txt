V19 - DRD Admin Panel Alignment + Induksi Achievement Fix

Perubahan utama:
1. Admin Panel DRD disamakan dengan pola aplikasi lain:
   - Form Buat User & Mapping Akses DRD
   - Upload Bulk Mapping Akses DRD via Excel
   - Row Data Mapping Akses DRD + Nonaktif
   - KPI ringkas dan shortcut operasional

2. Achievement Induksi diperbaiki:
   - Tidak lagi otomatis 100% ketika Masa Dinas Habis > 0 dan Induksi Closed = 0.
   - Formula: Induksi Closed / Total Wajib Induksi x 100%.
   - Total Wajib Induksi memperhitungkan masa dinas habis, open induksi, closed induksi, dan butuh input periode.

3. Tidak ada SQL baru wajib untuk v19.
   - Tetap jalankan SQL v17/v18 kalau belum pernah dijalankan.

Cara jalan:
npm install
npm run dev

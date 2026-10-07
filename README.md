# ZAIN.NET — Perapih Skripsi Otomatis V3

Versi **GitHub Pages upload langsung**. Tidak perlu npm, Vite, atau build lokal.

## Fokus perbaikan V3
- Cover dan semua halaman sebelum BAB I dilindungi dari pembersihan Enter massal.
- Cover 2, Halaman Persetujuan, Halaman Pengesahan, Abstrak, Kata Pengantar, Daftar Isi/Tabel/Gambar dipastikan mulai halaman baru memakai Page Break Before.
- BAB I–V dibersihkan dari Enter kosong di sekitarnya dan dipastikan mulai halaman baru.
- Sub BAB/Sub-sub BAB: Enter kosong yang menempel sebelum/sesudah judul dibersihkan, manual Shift+Enter pada judul dirapikan, Keep with next/Keep lines diterapkan.
- TOC tetap native Microsoft Word Level 1–3.
- Daftar Tabel/Gambar manual tidak didobel saat daftar dinamis dibuat.
- Tombol **Semua (Aman)** otomatis memakai Mode Aman.
- Format isi dalam Mode Aman tidak mengubah font run-level dan melewati format khusus/kutipan/list.

## Upload ke GitHub Pages
Upload semua file di folder ini langsung ke root repository:
`index.html`, `app.js`, `processor.js`, `styles.css`, `jszip.min.js`, `README.md`, `CARA_UPLOAD.txt`.

Lalu pilih **Settings → Pages → Deploy from a branch → main → /(root)**.

DASHBOARD MONITORING KUALITAS PRODUKSI ROKOK SKT (SIX SIGMA DMAI)

Isi folder
- app.py ............................ kode dashboard Streamlit (data Proyek Akhir sudah tertanam di dalamnya)
- requirements.txt .................. library yang dibutuhkan
- .streamlit/config.toml ............ tema warna (royal blue)
- template/Template_Data_SKT_kosong.xlsx  ... template untuk diisi perusahaan
- template/Template_Data_SKT_contoh.xlsx  ... template berisi data Proyek Akhir (contoh pengisian)

Cara kerja
1. Saat dibuka, dashboard langsung menampilkan data bawaan Proyek Akhir (66 subgrup, 30 Maret-30 Juni 2026).
   Folder data/ tidak diperlukan lagi.
2. Perusahaan mengunduh template (tombol di sidebar atau halaman "Data & Template"), mengisi sheet
   "Data Harian" dan "FMEA", lalu mengunggahnya lewat sidebar: Sumber data > Unggah data perusahaan.
3. Dashboard menghitung sendiri proporsi cacat, batas kendali (BKA/BKB per subgrup), peta P revisi,
   runs test, indeks kapabilitas, DPMO, level sigma, Pareto, dan RPN.
4. Peta P revisi dibentuk otomatis (titik di luar batas dikeluarkan berulang sampai terkendali) dan
   dapat diubah di sidebar > Pengaturan analisis. Workbook olah data format lama (sheet "Peta Kendali 1")
   juga masih bisa diunggah.

Diagram sebab-akibat (Fishbone) dan FMEA
- Analyze > Diagram Sebab-Akibat: Fishbone hasil penelitian (6 kategori, 18 penyebab; masalah utama di kepala ikan).
  Perusahaan dapat menambah, mengubah, atau menghapus penyebab lewat "Kelola penyebab". Perubahan berlaku selama sesi
  dan tidak mengubah data penelitian; tombol "Kembalikan" memulihkan 18 penyebab asli.
- Improve > "FMEA Penelitian (contoh pengisian)": FMEA asli penelitian, tetap dan tidak terpengaruh data unggahan.
- Improve > "FMEA Perusahaan (isi sendiri)": tabel isian (Faktor, Penyebab, Potensi, Dampak, S, O, D, Usulan);
  RPN dihitung otomatis dan dapat diunduh. Bila file unggahan memuat sheet FMEA, tabel ini terisi darinya.

Dasar data per analisis (sinkron dengan workbook Excel olah data)
- Peta P awal ................. 66 subgrup
- Peta P revisi ............... 54 subgrup (12 subgrup di luar batas dikeluarkan)
- Kapabilitas proses .......... 54 subgrup
- Runs test ................... 54 subgrup; proporsi = median tidak diikutsertakan (pembulatan 4 desimal)
- Level sigma ................. 66 subgrup, 1 opportunity (DPO = DPU)
- Diagram Pareto .............. 66 subgrup, tanpa ambang prioritas

Halaman: Ringkasan (jawaban tujuan penelitian) · Define (SIPOC) · Measure · Analyze (peta P, runs test,
kapabilitas, level sigma, Pareto, diagram sebab-akibat) · Improve (FMEA) · Data & Template.

Deploy ke Streamlit Community Cloud
1. Unggah app.py, requirements.txt, folder .streamlit, dan folder template ke repository GitHub.
2. Di Streamlit Community Cloud pilih repository tersebut, Main file path: app.py, lalu Deploy.

PENTING: data asli PT. XYZ tertanam di app.py. Jika repository GitHub bersifat publik, siapa pun dapat
melihat data tersebut. Gunakan repository privat atau pastikan perusahaan telah menyetujui.

Prinsip visualisasi (Cole Nussbaumer Knaflic, Storytelling with Data): judul berupa kesimpulan, warna
abu-abu sebagai dasar dengan satu aksen royal blue, merah hanya untuk titik di luar batas kendali,
label langsung pada data (tanpa legenda), tanpa gridline, tanpa pie/donut chart, tanpa sumbu ganda.

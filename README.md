# 🌴 RKP Monitor — Dashboard Monitoring Rencana Kerja Proyek

Dashboard Streamlit untuk memantau dan membandingkan file Excel Rencana Kerja Proyek (RKP)
lintas proyek: total biaya, biaya/Ha, target fisik per Catur Wulan, komposisi biaya per
pekerjaan, dan (jika sudah tersedia) perbandingan Rencana vs Realisasi.

## Menjalankan di komputer sendiri

```bash
pip install -r requirements.txt
streamlit run app.py
```

Lalu buka `http://localhost:8501`, upload file `.xlsx` RKP lewat panel kiri.

## Deploy online lewat GitHub + Streamlit Community Cloud (gratis)

**1. Buat repo GitHub baru**
- Buka https://github.com/new, beri nama repo (mis. `rkp-monitor`), pilih **Private** kalau
  data biaya proyek ini sensitif (disarankan), lalu klik **Create repository**.

**2. Upload isi folder ini ke repo**

Lewat terminal (ganti `<username>` dan `<repo>`):

```bash
cd rkp_streamlit
git init
git add .
git commit -m "Initial commit: RKP Monitor dashboard"
git branch -M main
git remote add origin https://github.com/<username>/<repo>.git
git push -u origin main
```

Atau lewat web: buka repo → **Add file → Upload files** → seret semua isi folder ini
(`app.py`, `requirements.txt`, folder `.streamlit/`, `README.md`) → **Commit changes**.

**3. Deploy di Streamlit Community Cloud**
- Buka https://share.streamlit.io, login pakai akun GitHub kamu.
- Klik **Create app** → **Deploy a public app from GitHub** (atau **From existing repo**).
- Pilih repo `<username>/<repo>`, branch `main`, main file path `app.py`.
- Klik **Deploy**. Tunggu 1-2 menit sampai app selesai build.
- Kamu akan dapat URL publik seperti `https://<nama-app>.streamlit.app`.

**Kalau repo private:** Streamlit Community Cloud tetap bisa deploy dari repo private
(perlu mengizinkan akses ke repo tersebut saat connect GitHub), tapi app yang sudah live
tetap bisa diakses lewat URL oleh siapa saja kecuali kamu aktifkan **App settings → Sharing →
Who can view this app** dan batasi ke email tertentu (fitur ini butuh akun Streamlit yang
sudah verifikasi org / viewer-based access, cek pengaturan terbaru di dashboard Streamlit
Cloud kamu).

## Update dashboard setelah ada perubahan

Setiap kali kamu edit `app.py` dan push ke GitHub (`git add . && git commit -m "update" && git push`),
Streamlit Community Cloud otomatis rebuild dan redeploy app dalam beberapa menit.

## Catatan penting: penyimpanan data

Dashboard ini **memproses file di memori sesi (session_state)** — begitu tab browser
ditutup atau di-refresh, data yang sudah diupload hilang dan perlu diupload ulang. Ini
sengaja dibuat sederhana (sesuai kebutuhan: upload manual tiap ada update RKP terbaru).

Kalau ke depan kamu ingin data tersimpan permanen dan bisa dilihat bersama tim tanpa upload
ulang tiap orang, opsi lanjutannya:
- Simpan hasil parsing ke Google Sheets / database (mis. Supabase, SQLite di volume
  persisten) setiap kali file diupload.
- Atau host di server sendiri (bukan Streamlit Community Cloud yang stateless) dengan disk
  permanen.

Kabari saya kalau mau saya bantu tambahkan salah satu opsi di atas.

## Format file yang didukung

Dashboard membaca sheet bernama **`RKP`** di tiap file `.xlsx`, dengan pola header:
`No. | Pekerjaan | Volume | Biaya | ... | TAHUN I / II / III... > Catur Wulan 1/2/3 > Fisik | Biaya`
dan baris **`GRAND TOTAL BIAYA`** di akhir tabel. Ini sudah disesuaikan dengan struktur
RKP yang kamu pakai saat ini.

Untuk data **Realisasi**: tambahkan sheet baru di file yang sama dengan nama mengandung kata
`Realisasi` (mis. `REALISASI` atau `Realisasi Biaya`), dengan struktur tabel yang sama persis
seperti sheet `RKP` (header Pekerjaan/Volume/Biaya/Fisik yang sama). Dashboard akan otomatis
mendeteksi dan menampilkan perbandingan Rencana vs Realisasi + % capaian.

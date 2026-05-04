# Dashboard Ekspor-Impor Indonesia — Deploy ke Railway

## Struktur File
```
├── app.py                  # Aplikasi utama Dash
├── requirements.txt        # Dependensi Python
├── Procfile                # Perintah start untuk Railway/Heroku
├── railway.toml            # Konfigurasi Railway
├── .gitignore
├── data_trademap.xlsx      # ⚠️  Upload manual (lihat langkah 3)
└── bop_indonesia.db        # ⚠️  Upload manual (lihat langkah 3)
```

---

## Cara Deploy ke Railway

### 1. Push ke GitHub
```bash
git init
git add app.py requirements.txt Procfile railway.toml .gitignore
git commit -m "initial deploy"
git remote add origin https://github.com/USERNAME/REPO.git
git push -u origin main
```

### 2. Buat proyek di Railway
1. Buka https://railway.app → **New Project** → **Deploy from GitHub repo**
2. Pilih repository Anda
3. Railway akan otomatis mendeteksi `Procfile` dan memulai build

### 3. Upload file data (data_trademap.xlsx & bop_indonesia.db)

**Opsi A — Commit ke repo (paling mudah, cocok untuk data ≤ 100 MB):**
```bash
git add data_trademap.xlsx bop_indonesia.db
git commit -m "add data files"
git push
```
> Hapus baris `*.db` dan `*.xlsx` dari `.gitignore` terlebih dahulu.

**Opsi B — Railway Volume (untuk data besar / sering berubah):**
1. Di Railway dashboard → proyek Anda → **+ Add Volume**
2. Mount path: `/data`
3. Upload file via Railway CLI:
   ```bash
   railway run -- bash -c "cp data_trademap.xlsx /data/ && cp bop_indonesia.db /data/"
   ```
4. Set environment variable `DATA_DIR=/data` di Railway dashboard

### 4. Set Environment Variables di Railway
Di Railway dashboard → proyek → **Variables**, tambahkan:

| Variable       | Value                                      | Keterangan                          |
|----------------|--------------------------------------------|-------------------------------------|
| `BPS_API_KEY`  | `c390bc3265694cce3a446082f9747178`         | API key BPS (wajib)                 |
| `DATA_DIR`     | `/data`                                    | Hanya jika pakai Volume (Opsi B)    |
| `BOP_DB_FILE`  | `bop_indonesia.db`                         | Nama file DB (opsional)             |
| `TM_XLSX_FILE` | `data_trademap.xlsx`                       | Nama file Excel (opsional)          |
| `MAX_WORKERS`  | `8`                                        | Thread pool (turunkan jika OOM)     |
| `CACHE_TTL`    | `600`                                      | Cache API dalam detik (opsional)    |

> `PORT` **tidak perlu** diset manual — Railway inject otomatis.

### 5. Verifikasi Deploy
Setelah build selesai, Railway akan memberi URL seperti:
`https://nama-proyek.up.railway.app`

Buka URL tersebut — dashboard siap digunakan.

---

## Troubleshooting

| Masalah | Solusi |
|---------|--------|
| Build gagal "No module named X" | Pastikan `requirements.txt` lengkap |
| App crash saat start | Cek log Railway, pastikan `server = app.server` ada di `app.py` |
| Data BPS tidak muncul | Periksa `BPS_API_KEY` di Variables |
| Tab SEKI error merah | File `bop_indonesia.db` belum terupload / path salah |
| Tab Mirroring error | File `data_trademap.xlsx` belum terupload / path salah |
| OOM / timeout | Kurangi `MAX_WORKERS` ke 4-6, atau upgrade Railway plan |

---

## Pengembangan Lokal
```bash
pip install -r requirements.txt
# Letakkan data_trademap.xlsx dan bop_indonesia.db di folder yang sama dengan app.py
python app.py
# Buka http://localhost:8050
```

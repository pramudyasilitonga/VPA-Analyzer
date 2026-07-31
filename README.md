# VPA Analyzer

Skrip otomatisasi DevOps untuk menganalisis rekomendasi **Vertical Pod Autoscaler (VPA)** dari ratusan service Kubernetes, lalu mengekspor hasilnya ke file Excel.

## Fitur

- Membaca data dari `sample_vpa.json` (mode testing) atau `kubectl get vpa -A -o json` (mode live)
- Loop aman pada semua item VPA; service tanpa rekomendasi di-skip dengan log info
- Menghitung request = target VPA, limit = target × 1.2 (20%)
- Ekspor ke `vpa_recommendations.xlsx` dengan kolom lengkap

## Prasyarat

- Python 3.10+
- Untuk mode live: `kubectl` terkonfigurasi dan akses ke cluster Kubernetes

## Instalasi

```bash
pip install -r requirements.txt
```

## Menjalankan Skrip

### Mode testing (menggunakan sample lokal)

Pastikan file `sample_vpa.json` ada di folder yang sama, lalu jalankan:

```bash
python vpa_analyzer.py
```

### Mode live (kubectl)

Hapus atau rename file `sample_vpa.json`, lalu jalankan:

```bash
python vpa_analyzer.py
```

Skrip akan otomatis menjalankan:

```bash
kubectl get vpa -A -o json
```

## Output

File Excel: **`vpa_recommendations.xlsx`**

| Kolom | Deskripsi |
|-------|-----------|
| Namespace | Namespace Kubernetes |
| Nama Service | Nama resource VPA |
| Target CPU VPA | Rekomendasi CPU dari VPA |
| Target Memori VPA | Rekomendasi memory dari VPA |
| Rekomendasi Request CPU | Sama dengan target CPU |
| Rekomendasi Limit CPU | Target CPU × 1.2 |
| Rekomendasi Request Memori | Sama dengan target memory |
| Rekomendasi Limit Memori | Target memory × 1.2 |

## Konfigurasi (Environment Variables)

| Variable | Default | Deskripsi |
|----------|---------|-----------|
| `VPA_SAMPLE_FILE` | `sample_vpa.json` | Path file JSON sample untuk testing |
| `VPA_OUTPUT_FILE` | `vpa_recommendations.xlsx` | Nama file Excel output |
| `VPA_KUBECTL_CMD` | `kubectl get vpa -A -o json` | Perintah kubectl untuk mengambil data VPA |
| `VPA_LIMIT_MULTIPLIER` | `1.2` | Pengali untuk menghitung limit (20% di atas target) |
| `VPA_LOG_LEVEL` | `INFO` | Level logging (`DEBUG`, `INFO`, `WARNING`, `ERROR`) |

### Contoh penggunaan custom config

**Windows (PowerShell):**

```powershell
$env:VPA_OUTPUT_FILE = "hasil_vpa.xlsx"
$env:VPA_LIMIT_MULTIPLIER = "1.5"
python vpa_analyzer.py
```

**Linux / macOS:**

```bash
export VPA_OUTPUT_FILE=hasil_vpa.xlsx
export VPA_LIMIT_MULTIPLIER=1.5
python vpa_analyzer.py
```

## Catatan

- Jika VPA memiliki beberapa container, target CPU dan memory dijumlahkan dari semua `containerRecommendations`.
- Service yang rekomendasinya kosong atau belum tersedia akan di-skip tanpa menghentikan skrip.

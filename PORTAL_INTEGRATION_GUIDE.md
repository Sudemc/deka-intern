# DEKA BOM Karşılaştırma Modülü - Şirket Portalı Entegrasyon Kılavuzu

Bu dal (`feature/portal-bom-comparison`), projedeki PDF/OCR ve yapay zekâ kütüphanelerinden tamamen arındırılmış, yalnızca **BOM Karşılaştırma Motoru (Multi-BOM & ERP Comparison)** için hazırlanmış hafif (lightweight) ve bağımsız sürümdür.

---

## 1. Paket ve Bağımlılıklar

Modül tamamen çevrimdışı (offline) çalışır, dış API anahtarı veya internet bağlantısı gerektirmez.

Gereken kütüphaneler:
```bash
pip install -r requirements.txt
# veya
pip install pandas openpyxl fastapi uvicorn python-multipart
```

---

## 2. Klasör ve Dosya Yapısı

Portala aktarılması gereken çekirdek dosyalar:

```text
src/bom_enricher/
├── comparison.py       # Ana karşılaştırma motoru (Multi-Pass Matching, Multi-BOM desteği)
├── semantic_parser.py  # Elektriksel birim normalizasyonu (uF, nF, mOhm, kOhm, SI çevrimi)
├── excel_io.py         # Müşteri/ERP Excel okuma ve akıllı başlık tespit mekanizması
├── models.py           # Veri modelleri (CustomerBomRow, PartRecord vb.)
├── web.py              # FastAPI REST API servisi ve indirme yönetimi
└── static/             # Test / demo için basit web arayüzü
```

---

## 3. Kullanım Yolları

### Yöntem A: Doğrudan Python Fonksiyonu Olarak Çağırma (En Kolay Yol)
Portalınızın backend'i Python (Django / Flask / FastAPI / Celery) ise doğrudan modülü import edip çalıştırabilirsiniz:

```python
from pathlib import Path
from bom_enricher.comparison import compare_workbooks, load_rulebook

# 1. Tek veya birden fazla Müşteri BOM Excel'i
customer_files = [
    Path("yuklenenler/StopTail_BOM.xlsx"),
    Path("yuklenenler/TurnRev_BOM.xlsx"),
]

# 2. ERP Malzeme Listesi Excel'i
erp_file = Path("yuklenenler/ERP_BOM.xlsx")

# 3. İsteğe Bağlı Kural Dosyası (Malzeme Tipleri Sözlüğü)
rules_file = Path("data/Malzeme_Tipleri.xlsx")  # Opsiyoneldir, verilmezse kuralsız semantik kıyaslama yapar

# 4. Çıktı Rapor Yolu
output_report_path = Path("cikti/Kiyaslama_Raporu.xlsx")

# Karşılaştırmayı çalıştır:
result = compare_workbooks(
    customer_path=customer_files,   # Path veya list[Path]
    erp_path=erp_file,
    rules_path=rules_file,
    output_path=output_report_path  # Excel raporu buraya yazılır
)

print(f"Toplam Satır: {len(result.rows)}")
for row in result.rows:
    print(row["Durum"], row["Part Number"], row["Kaynak Kart / BOM"], row["Hata/Uyarı Durumu"])
```

---

### Yöntem B: REST API Servisi Olarak Kullanma (FastAPI)
Eğer portal farklı bir dille (C#, Java, Node.js, PHP vb.) yazılmışsa, bu servisi bağımsız bir mikroservis / Docker container olarak ayağa kaldırabilirsiniz:

```bash
uvicorn bom_enricher.web:app --host 0.0.0.0 --port 8000
```

#### API Endpoint: `POST /api/compare`
* **Girdiler (`multipart/form-data`):**
  * `customer_files`: Birden fazla `.xlsx` dosyası (veya tek dosya için `customer_file`)
  * `erp_file`: ERP `.xlsx` dosyası (zorunlu)
  * `rules_file`: Malzeme tipleri `.xlsx` (opsiyonel)
* **Dönen JSON Yanıtı:**
```json
{
  "result_id": "abc12345...",
  "row_count": 150,
  "error_count": 2,
  "warning_count": 5,
  "clean_count": 143,
  "download_url": "/api/results/abc12345.../workbook",
  "rows": [
    {
      "Durum": "UYGUN",
      "Müşteri Part Number": "RC0603FR-0710KL",
      "ERP Part Number": "RC0603FR-0710KL",
      "Part Number": "RC0603FR-0710KL",
      "Müşteri Qty": "8",
      "ERP Qty": "8",
      "Müşteri Value": "10k 1%",
      "ERP Description": "SMD 0603, 10k, 1%",
      "Müşteri TH/SMD": "SMD",
      "ERP TH/SMD": "SMD",
      "Kaynak Kart / BOM": "StopTail.xlsx (5 adet), TurnRev.xlsx (3 adet)",
      "Müşteri Referans": "ST1_R1, ST1_R2, TR1_R5",
      "ERP Referans": "ST1_R1, ST1_R2, TR1_R5",
      "Hata/Uyarı Durumu": ""
    }
  ]
}
```

---

## 4. Çıktı Raporu Formatı (13 Sütun)

Oluşturulan Excel dosyasındaki sütunlar:

| No | Sütun Adı | Açıklama |
|---|---|---|
| 1 | **Durum** | `HATA` (kırmızı), `UYARI` (sarı) veya `UYGUN` (temiz) |
| 2 | **Müşteri Part Number** | Müşterinin yazdığı parça kodu |
| 3 | **ERP Part Number** | ERP'de eşleşen parça kodu |
| 4 | **Müşteri Qty** | Müşteri toplam adedi (sayısal) |
| 5 | **ERP Qty** | ERP toplam adedi (sayısal) |
| 6 | **Müşteri Value** | Müşteri teknik değeri / açıklaması |
| 7 | **ERP Description** | ERP malzeme açıklaması |
| 8 | **Müşteri TH/SMD** | Müşteri montaj tipi (`SMD` veya `TH`) |
| 9 | **ERP TH/SMD** | ERP montaj tipi (`SMD` veya `TH`) |
| 10 | **Kaynak Kart / BOM** | Parçanın geldiği dosya ve çoklu BOM varsa kart başına adet kırılımı |
| 11 | **Müşteri Referans** | Müşteri devre pozisyonları (`C1, C2`, `TR1_C5` vb.) |
| 12 | **ERP Referans** | ERP devre pozisyonları |
| 13 | **Hata/Uyarı Durumu** | Tespit edilen uyumsuzluklar ve açık referans farkı |

---

## 5. Çoklu BOM ve Referans Eşleme Mantığı

1. **Otomatik Önek Eşleme:** ERP referanslarında `TR1_C1`, `ST1.C1` veya `FM-R5` gibi kart önekleri varsa, yüklenen müşteri dosyalarının adları veya sekme adlarıyla (örn. `StopTail1.xlsx` -> `ST1`) otomatik eşleştirilir.
2. **Kırılım ve Toplam:** Farklı kartlarda aynı parçadan kullanılıyorsa (`TR1`'de 3 adet, `ST1`'de 5 adet), `Müşteri Qty` alanı 8 kalır, `Kaynak Kart / BOM` alanında ise `StopTail1: 5 adet, TurnRev1: 3 adet` şeklinde detay görünür.
3. **Anında Fark (Diff) Gösterimi:** Referanslarda uyuşmazlık olduğunda hata hücresinde tam liste korunurken, doğrudan farklı olan elemanlar gösterilir:  
   `Dizgi/Referans farkı (Fark: Müşteride Farklı/Fazla: ST3_R56 | ERP'de Farklı/Fazla: ST3_R55; Müşteri: ..., ERP: ...)`

---

## 6. Testleri Çalıştırma

Tüm kontrolleri doğrulamak için:
```bash
pytest tests/
```

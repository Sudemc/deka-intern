# DEKA Elektronik — BOM Karşılaştırma & Veri Doğrulama Motoru

<div align="center">
  <img src="src/bom_enricher/static/deka-logo.png" alt="DEKA Elektronik Logo" width="240">
  <br><br>
  <p><strong>Müşteri BOM listeleri ile ERP üretim kayıtlarını karşılaştıran, doğrulayan ve uyuşmazlıkları raporlayan %100 çevrimdışı (offline) mühendislik motoru.</strong></p>
</div>

---

> [!IMPORTANT]
> **Şirket Portalı (ERP Yönetim Sistemi) Entegrasyonu İçin:**  
> Bu modülü DEKA Şirket Portalı'na entegre edecek veya bağımsız mikroservis olarak çalıştıracak yazılım mühendisleri için hazırlanan kapsamlı kılavuz:  
> 📖 **[Şirket Portalı Entegrasyon ve Kurulum Kılavuzu (`PORTAL_INTEGRATION_GUIDE.md`)](PORTAL_INTEGRATION_GUIDE.md)**

---

## 📌 Hangi Daldan Çalışmalısınız? (Branch Rehberi)

| Dal (Branch) | Kimler İçin? | Kapsam |
| :--- | :--- | :--- |
| **`portal-bom-comparison`** *(veya `feature/portal-bom-comparison`)* | **Şirket Portalı & Backend Mühendisleri** | Ağır OCR, PDF ve yapay zekâ bağımlılıklarından arındırılmış; yalnızca **Zero-Defect BOM Karşılaştırma Motoru**, semantik ayrıştırıcı ve hafif FastAPI web servisini içeren bağımsız sürüm. |
| **`main` / `dev-sudem`** | **Geliştiriciler & Tam Sistem** | BOM Karşılaştırma Motoru + PDF Datasheet Analizi & ERP Sözlük Eşleme modüllerinin tamamını içeren ana geliştirme dalı. |

---

## ✨ Temel Özellikler ve Çekirdek Domain Mantığı

* **3 Aşamalı Akıllı Eşleştirme Motoru:**
  1. **Aşama 1 (MPN / Part Number):** Normalizasyon, alternatif üretici kodları, birleşik kodlar (`4E12000022(AC0603FR-0710KL)`) ve parantez varyantları üzerinden tam eşleşme.
  2. **Aşama 2 (RefDes / Devre Pozisyonu):** Parça kodu farklı olsa dahi devre referansları (`C1, R12, U3`) üzerinden (%100 ve kesişim) akıllı eşleştirme ve alternatif parça tespiti.
  3. **Aşama 3 (Açıklama / Description Fallback):** Kodu bulunmayan kalemlerin elektriksel açıklama ve değer üzerinden eşleştirilmesi.
* **Evrensel Elektriksel Semantik Ayrıştırıcı (`semantic_parser.py`):**
  * Kapasitans (`100nF == 0.1uF`), direnç (`4K7 == 4.7kOhm`, `1 MOhms == 1M`), indüktans, voltaj, tolerans, akım, güç, frekans, sıcaklık, kılıf (`0402`, `0603`, `SOIC-8`, `QFN-32` vb.), montaj tipi (`SMD` / `TH`), kutupluluk ve renk parametrelerini matematiksel olarak normalize eder.
* **Otomotiv Simetrisi ve L/R (Sol / Sağ) Çiftleme Algoritması:**  
  Otomotiv lambalarında sol ve sağ için tek satırda $2\times$ referans yazıldığında motor simetriyi tanır, sahte mükerrerlik hatası üretmez ve ERP ile doğru miktar uzlaştırması yapar.
* **Uydu Kart (Satellite Board) Önek Eşleme:**  
  Tek kartta çift kanal barındıran alt kartlar (`ST2` -> `ST2L` + `ST2R`) ve birleşik ERP dosyalarındaki kart önekleri otomatik eşleştirilir.
* **Aktif Montaj (SMD) ve Dizilmez (DNP / Process: NM) Otomatik Ayrıştırması:**  
  `DNP`, `NM`, `DİZİLMEZ`, `FİDUCIAL`, `TP` (Test Point) ve mekanik/LOGO kalemlerini otomatik algılar; aktif montaj satırlarıyla birleştirilmiş ERP kayıtlarını referans kümeleri üzerinden hatasız ayrıştırır.
* **13 Sütunlu Kurumsal Excel Raporu (`openpyxl`):**  
  * 🔴 **Açık Kırmızı (`HATA`):** Gerçek donanım, değer, kılıf veya miktar kusurları.
  * 🟡 **Sarı (`UYARI`):** Alternatif parça kullanımı, RefDes yazım farkları ve teknik uyarılar.
  * 🔵 **Açık Mavi (`UYARI - Montaj Dışı`):** `DNP / Process: NM` (Dizilmez) kalemler ve ERP'ye girilmeyen montaj dışı unsurlar.
  * ⚪ **Temiz Hücre (`UYGUN`):** %100 tam eşleşen üretim onaylı kalemler (`Hata/Uyarı Durumu` hücresi tamamen boş bırakılır).

---

## 📂 Proje Dizin Yapısı

```text
├── data/
│   └── rules.xlsx                  # DEKA ERP Malzeme Tipleri ve SORTNO kural sözlüğü
├── src/
│   └── bom_enricher/
│       ├── __init__.py             # Paket dışa aktarımları (compare_workbooks, RuleBook vb.)
│       ├── comparison.py           # 3 aşamalı BOM karşılaştırma motoru ve Excel rapor üreticisi
│       ├── excel_io.py             # Çoklu sekme/dosya ve CSV/XLSX akıllı tablo okuyucu
│       ├── semantic_parser.py      # Elektriksel değer, birim ve kılıf normalizasyon motoru
│       ├── models.py               # Veri modelleri (CustomerBomRow, ComponentData vb.)
│       ├── web.py                  # FastAPI REST API ve web sunucusu (/api/compare, /api/health)
│       └── static/                 # Kurumsal Türkçe Web Arayüzü (index.html, logolar)
├── tests/                          # Otomatik birim ve uçtan uca doğrulama testleri (pytest)
├── PORTAL_INTEGRATION_GUIDE.md     # Şirket portalı entegrasyon ve API kılavuzu
├── pyproject.toml                  # Python paket yapılandırması
├── requirements.txt                # Bağımlılık listesi
└── baslat.bat                      # Windows tek tıkla başlatma betiği
```

---

## 🚀 Kendi Bilgisayarında Kurulum (Local Setup)

### 1. Repoyu Klonlayın ve Dala Geçin
```bash
git clone https://github.com/Sudemc/deka-intern.git
cd deka-intern

# Şirket portalı BOM karşılaştırma dalı için:
git checkout portal-bom-comparison
```

### 2. Sanal Ortamı Kurun ve Bağımlılıkları Yükleyin

**Windows (PowerShell):**
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
pip install -e .
```

**macOS / Linux:**
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
pip install -e .
```

### 3. Testleri Çalıştırın
```bash
pytest
```

### 4. Web Arayüzünü Başlatın
**Windows:** Çift tıklayarak `baslat.bat` dosyasını çalıştırabilirsiniz.  
**Terminal:**
```bash
uvicorn bom_enricher.web:app --host 0.0.0.0 --port 8000 --reload
```
Tarayıcınızdan **`http://127.0.0.1:8000`** adresine giderek karşılaştırmayı yapabilir ve 13 sütunlu renkli Excel raporunu indirebilirsiniz.

---

## 🔌 Şirket Portalı Entegrasyon Özeti

### Yöntem 1: Doğrudan Python Kodu Olarak Çağırma (En Hızlı)
```python
from pathlib import Path
from bom_enricher import compare_workbooks

result = compare_workbooks(
    customer_path=[Path("Musteri_BOM.xlsx")],
    erp_path=Path("ERP_BOM.xlsx"),
    output_path=Path("Karsilastirma_Raporu.xlsx"),  # 13 sütunlu renkli Excel çıktısı
)

print(f"Toplam Satır: {len(result.rows)} | Toplam Parça: {result.total_part_count}")
for row in result.rows:
    print(
        row["Durum"],
        row["Müşteri Part Number"],
        row["ERP Part Number"],
        row["Müşteri Qty"],
        row["ERP Qty"],
        row["Hata/Uyarı Durumu"],
    )
```

### Yöntem 2: REST API / Docker Mikroservisi
* **Endpoint:** `POST /api/compare` (`multipart/form-data`)
* **Parametreler:** `customer_files` (`.xlsx` veya `.csv`), `erp_file` (`.xlsx` veya `.csv`), `rules_file` (opsiyonel `.xlsx`)
* **Dönen Yanıt:** JSON (özet metrikler, satır detayları ve `/api/results/{result_id}/workbook` indirme adresi).

> Ayrıntılı API sözleşmesi, Dockerfile ve entegrasyon örnekleri için **[`PORTAL_INTEGRATION_GUIDE.md`](PORTAL_INTEGRATION_GUIDE.md)** dosyasına bakınız.

---

<div align="center">
  <sub>DEKA Elektronik © 2026</sub>
</div>

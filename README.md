# DEKA Elektronik — BOM Karşılaştırma & Analiz Motoru

<div align="center">
  <img src="src/bom_enricher/static/deka-logo.png" alt="DEKA Elektronik Logo" width="240">
  <br><br>
  <p><strong>Müşteri BOM listeleri ile ERP üretim kayıtlarını karşılaştıran, doğrulayan ve uyuşmazlıkları raporlayan %100 çevrimdışı motor.</strong></p>
</div>

---

> [!IMPORTANT]
> **Şirket Portalı (ERP Yönetim Sistemi) Entegrasyonu İçin:**  
> Bu modülü DEKA Şirket Portalı'na eklemek veya bağımsız mikroservis olarak çalıştırmak isteyen yazılım mühendisleri için hazırlanan kapsamlı entegrasyon rehberi:  
> 📖 **[Şirket Portalı Entegrasyon ve Kurulum Kılavuzu (PORTAL_INTEGRATION_GUIDE.md)](PORTAL_INTEGRATION_GUIDE.md)**

---

## 📌 Hangi Daldan Çalışmalısınız? (Branch Rehberi)

| Dal (Branch) | Kimler İçin? | Kapsam |
| :--- | :--- | :--- |
| **`feature/portal-bom-comparison`** | **Şirket Portalı & Backend Mühendisleri** | Ağır OCR, PDF ve yapay zekâ bağımlılıklarından arındırılmış; sadece **BOM Karşılaştırma Motoru** ve hafif FastAPI web servisini içeren bağımsız, hafif sürüm. |
| **`dev-sudem`** | **Geliştiriciler & Tam Sistem** | Datasheet Analizi (VLM / LLM) ve BOM Karşılaştırma dahil tüm modülleri içeren ana geliştirme dalı. |

---

## ✨ Temel Özellikler ve Çekirdek Domain Mantığı

* **3 Aşamalı Akıllı Eşleştirme Motoru:**
  1. **Aşama 1 (MPN / Part Number):** Normalizasyon, alternatif üretici kodları ve parantez varyantları üzerinden tam eşleşme.
  2. **Aşama 2 (RefDes / Devre Pozisyonu):** Parça kodu farklı olsa dahi devre referansları üzerinden (%100 ve kesişim) akıllı eşleştirme.
  3. **Aşama 3 (Açıklama / Description Fallback):** Kodu bulunmayan kalemlerin elektriksel açıklama üzerinden eşleştirilmesi.
* **Otomotiv Simetrisi ve L/R (Sol / Sağ) Çiftleme Algoritması:**  
  Otomotiv lambalarında sol ve sağ için tek satırda $2\times$ referans yazıldığında motor simetriyi tanır, mükerrerlik hatası üretmez ve ERP ile doğru miktar uzlaştırması yapar.
* **Uydu Kart (Satellite Board) Önek Eşleme:**  
  Tek kartta çift kanal barındıran alt kartlar (`ST2` -> `ST2L` + `ST2R`) ERP ile otomatik eşleştirilir.
* **Aktif (SMD) ve Dizilmez (Process: NM) Otomatik Ayrıştırması:**  
  Müşterinin aktif (30 adet) ve dizilmez (16 adet) olarak ayırdığı parçaları ERP tek kalemde 46 girmiş olsa dahi motor aktif kısmı `UYGUN`, dizilmez kısmı ise `UYARI` ("Process: NM - Dizilmemelidir") olarak ayırır; sahte miktar uyuşmazlığı üretmez.
* **LED Lot/BIN ve Alternatif Kod Temizliği:**  
  ERP hücrelerindeki `BIN: ...` gibi ekler ayıklanır; `4E12000022(AC0603FR-0710KL)` formatındaki birleşik kodlar çift yönlü çözümlenir.
* **Açık Referans Farkı (RefDes Diff):**  
  Referans uyuşmazlıklarında doğrudan hangi referansların ERP'de fazla veya müşteride eksik olduğu hücre içinde açıkça gösterilir.
* **13 Sütunlu Standart Excel Raporu:**  
  * 🔴 **Açık Kırmızı (`HATA`):** Gerçek donanım ve miktar kusurları.
  * 🔵 **Açık Mavi (`UYARI - Montaj Dışı`):** `Process: NM` (Dizilmez) kalemler, Ham PCB.
  * 🟡 **Sarı (`UYARI`):** RefDes yazım farkları ve teknik uyarılar.
  * ⚪ **Beyaz / Yeşil (`UYGUN`):** %100 tam eşleşen üretim onaylı kalemler.

---

## 🚀 Kendi Bilgisayarında Kurulum (Local Setup)

Projeyi yerelinizde çalıştırmak için aşağıdaki adımları izleyin:

### 1. Repoyu Klonlayın ve Dala Geçin
```bash
git clone https://github.com/deka-internship/erp-data-manager.git
cd erp-data-manager

# Sadece BOM karşılaştırma motoru ve portal entegrasyonu için:
git checkout feature/portal-bom-comparison
# veya tam geliştirme dalı için:
# git checkout dev-sudem
```

### 2. Sanal Ortamı Kurun ve Bağımlılıkları Yükleyin

**Windows (PowerShell):**
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

**macOS / Linux:**
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 3. Testleri Çalıştırın
```bash
pytest tests/
```
Tüm testlerin `passed` olduğunu doğrulayın (58+ test).

### 4. Web Arayüzünü Başlatın
**Windows:** Çift tıklayarak `baslat.bat` çalıştırabilirsiniz.  
**Terminal:**
```bash
uvicorn bom_enricher.web:app --host 0.0.0.0 --port 8000 --reload
```
Tarayıcınızdan **`http://127.0.0.1:8000`** adresine giderek karşılaştırmayı yapabilir veya sonuç Excel raporunu indirebilirsiniz.

---

## 🔌 Şirket Portalı Entegrasyon Özeti

Portal backend mühendisleri için iki temel kullanım yöntemi:

### Yöntem 1: Python Kodu Olarak Çağırma (En Kolay)
```python
from pathlib import Path
from bom_enricher.comparison import compare_workbooks

result = compare_workbooks(
    customer_path=[Path("StopTail.xlsx"), Path("Turn.xlsx")],
    erp_path=Path("ERP_BOM.xlsx"),
    output_path=Path("Karsilastirma_Raporu.xlsx")  # 13 sütunlu renkli Excel çıktısı
)

print(f"Toplam Satır: {len(result.rows)}")
for row in result.rows:
    print(row["Durum"], row["Part Number"], row["Müşteri Qty"], row["ERP Qty"], row["Hata/Uyarı Durumu"])
```

### Yöntem 2: REST API / Docker Mikroservisi
* **Endpoint:** `POST /api/compare` (`multipart/form-data`)
* **Parametreler:** `customer_files` (Excel dosyaları), `erp_file` (ERP Excel'i)
* **Dönen Yanıt:** JSON (istatistikler, satır detayları ve tek tıkla Excel indirme adresi `/api/results/{id}/workbook`).

> Ayrıntılı API sözleşmesi, Dockerfile ve kod örnekleri için [`PORTAL_INTEGRATION_GUIDE.md`](PORTAL_INTEGRATION_GUIDE.md) dosyasına bakınız.

---

<div align="center">
  <sub>DEKA Elektronik © 2026</sub>
</div>

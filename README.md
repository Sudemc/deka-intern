# DEKA Elektronik — BOM Karşılaştırma Motoru

<div align="center">
  <img src="src/bom_enricher/static/deka-logo.png" alt="DEKA Elektronik Logo" width="240">
  <br><br>
  <p><strong>Müşteri BOM listeleri ile ERP kayıtlarını karşılaştıran ve doğrulayan çevrimdışı analiz aracı.</strong></p>
</div>

---

## 📌 Genel Bakış

Bu proje, müşteri tarafından sağlanan Excel BOM (Bill of Materials) listeleri ile ERP malzeme kayıtları arasındaki uyuşmazlıkları tespit etmek için geliştirilmiştir. Sistem **%100 çevrimdışı (offline)** çalışır ve harici API bağımlılığı bulunmaz.

---

## ✨ Temel Özellikler

* **3 Aşamalı Akıllı Eşleştirme:**
  1. **MPN (Part Number):** Boşluk, noktalama ve ambalaj ekleri (`C/D`) yok sayılarak eşleştirilir.
  2. **RefDes (Devre Pozisyonu):** Parça kodu farklı olsa bile devre referansı (`C1, R1`) üzerinden eşleştirilir.
  3. **Açıklama (Fallback):** Kodu bulunmayan parçalar birebir açıklama üzerinden bağlanır.
* **Çoklu ERP Kodu Desteği:** Bir hücrede birden fazla alternatif kod olduğunda parçalar otomatik ayrıştırılır.
* **Teknik Değer Karşılaştırması:** Elektriksel birimler (`100nF` = `0.1uF`, `10k` = `10 Kohm`) ve montaj tipleri (`TH` / `SMD`) otomatik doğrulanır.
* **Koşullu Renklendirmeli Excel Raporu:**
  * ⚪ **Normal (Uygun):** Tam eşleşen sorunsuz satırlar.
  * 🟡 **Sarı (Uyarı):** Eksik kod veya bilgilendirici durumlar.
  * 🔴 **Kırmızı (Hata):** Miktar uyuşmazlığı ve teknik çelişkiler.
* **Opsiyonel Kural Sözlüğü:** Sistem sözlük olmadan da bağımsız olarak çalışır.

---

## 🚀 Kurulum ve Çalıştırma

Projeyi kendi bilgisayarınızda çalıştırmak için aşağıdaki adımları izleyin:

### 1. Repoyu Klonlayın ve Dala Geçin
```bash
git clone https://github.com/deka-internship/erp-data-manager.git
cd erp-data-manager
git checkout dev-sudem
```

### 2. Sanal Ortamı Kurun ve Bağımlılıkları Yükleyin

**Windows (PowerShell):**
```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -e ".[dev]"
```

**macOS / Linux:**
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -e ".[dev]"
```

### 3. Web Arayüzünü Başlatın
```bash
python -m bom_enricher.web
```

Tarayıcınızdan **`http://127.0.0.1:8000`** adresine gidin:
1. **Müşteri BOM** ve **ERP** Excel (`.xlsx`) dosyalarını seçin.
2. **Karşılaştırmayı Başlat** butonuna basın.
3. Sonuçları anlık inceleyin veya **Excel Raporunu İndir** ile tam raporu alın.

---

## 🧪 Testler

```powershell
.\.venv\Scripts\python.exe -m pytest -v
```

Tüm modüller ve eşleştirme kuralları test kapsamındadır (61 test).

---

<div align="center">
  <sub>DEKA Elektronik © 2026</sub>
</div>

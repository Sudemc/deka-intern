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

## 🚀 Hızlı Başlangıç

### 1. Kurulum

```powershell
# Sanal ortam oluşturma ve bağımlılıkları yükleme
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -e ".[dev]"
```

### 2. Web Arayüzünü Çalıştırma

```powershell
.\.venv\Scripts\python.exe -m bom_enricher.web
```

Tarayıcınızda açın: **`http://127.0.0.1:8000`**

1. Müşteri BOM ve ERP dosyalarını yükleyin.
2. **Karşılaştırmayı Başlat** butonuna basın.
3. Sonuçları görüntüleyin ve Excel raporunu indirin.

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

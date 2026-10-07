# DEKA Elektronik — ERP Data Manager & BOM Karşılaştırma Motoru

<div align="center">
  <img src="src/bom_enricher/static/deka-logo.png" alt="DEKA Elektronik Logo" width="280">
  <br><br>
  <p><strong>Müşteri BOM Listeleri ile ERP Kayıtlarını Karşılaştıran, Doğrulayan ve Raporlayan Akıllı Analiz Motoru</strong></p>

  [![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)](https://python.org)
  [![FastAPI](https://img.shields.io/badge/FastAPI-0.115%2B-009688?logo=fastapi&logoColor=white)](https://fastapi.tiangolo.com)
  [![Tests](https://img.shields.io/badge/Tests-61%20Passed-brightgreen?logo=pytest&logoColor=white)](#testler)
  [![Offline](https://img.shields.io/badge/Privacy-100%25%20Offline-success)](https://dekaelectronic.com)
</div>

---

## 📌 Proje Hakkında

**DEKA ERP Data Manager**, elektronik imalat ve tasarım süreçlerinde müşteri tarafından sağlanan **BOM (Bill of Materials)** çalışma kitapları ile dahili **ERP malzeme kayıtları** arasındaki uyuşmazlıkları tespit etmek üzere geliştirilmiş yüksek performanslı bir karşılaştırma motorudur.

Sistem; parça numarası formatlarındaki varyasyonları, alternatif eşdeğer kodları, devre referans (RefDes) dizilimlerini, elektriksel değer denkliklerini ve paket/kılıf bilgilerini akıllı algoritmalarla eşleştirir. **Tüm operasyon %100 çevrimdışı (offline) çalışır;** hiçbir veri dışarıya aktarılmaz.

---

## ✨ Temel Yetenekler

### 1. 🎯 Üç Aşamalı Akıllı Eşleştirme (Three-Pass Matching)
* **1. Aşama (MPN Eşleme):** Boşluk, noktalama (`-`, `_`, `/`), parantezli ambalaj varyantları (örn. `CL10B104KB8NNN(C/D)`) ve büyük/küçük harf farklarını yok sayarak birincil üretici kodlarını eşler.
* **2. Aşama (RefDes / Devre Dizilim Eşleme):** Part Number farklı veya alternatif olsa dahi (`C1, C2` vb.) devre referans dizilimini inceleyerek pozisyon bazlı eşleştirme sağlar.
* **3. Aşama (Açıklama / Fallback Eşleme):** Müşteri tarafında parça kodu belirtilmemiş satırları ERP açıklamasıyla birebir anlamsal eşleştirmeye tabi tutar.

### 2. ⚡ Çoklu Kod & Kardeş Kod (Sibling) Ayrıştırma
* ERP hücrelerinde boşlukla yazılmış çoklu alternatif üretici kodlarını (`CL21B104KBCNNNC 885012207098 CC0805KRX7R9BB104`) bağımsız olarak genişletir.
* Kodlardan biri müşteri listesinde eşleştiğinde, diğer kardeş alternatifleri raporda sahte hata ("Sistemde Yok") olarak üretmez.

### 3. 🔍 Semantik Elektriksel ve Teknik Değer Doğrulaması
* Birim normalizasyonu uygular:  
  * `100nF` = `0.1uF` = `100000pF`
  * `10k` = `10 Kohm` = `10000 Ohm`
  * `50V` = `50 Volt`
* Eşdeğer teknik değerler için yanlış alarm üretmez; yalnızca gerçek voltaj, tolerans ve paket (örn. `0805` vs `1206`) çelişkilerini **HATA** olarak yakalar.

### 4. 📐 Montaj Tipi (TH / SMD) ve Miktar (Qty) Denetimi
* Montaj sütunu boş olduğunda, bileşenin açıklamasından (`TH`, `SMD`, `THT`) montaj tipini otomatik olarak çıkarır.
* Müşteri Qty ile ERP Qty sayısal farklarını ve referans adedi (`C1, C2, C3` = 3 adet) uyuşmazlıklarını denetler.

### 5. 📊 Koşullu Renklendirmeli Master Excel Raporlama
* Yalnızca hatalı kalemleri değil; **doğrulanan tüm doğru satırları da** rapora dahil eder:
  * ⚪ **Normal / Beyaz:** Tam eşleşen, sorunsuz onaylanmış satırlar (`UYGUN`).
  * 🟡 **Açık Sarı:** Bilgilendirici uyarılar ve eksik kodlar (`UYARI`).
  * 🔴 **Açık Kırmızı:** Miktar veya teknik kılıf/voltaj çelişkileri (`HATA`).

### 6. 🛡️ Tamamen Opsiyonel Kural Sözlüğü
* Malzeme sözlüğü (`rules.xlsx`) sisteme opsiyonel olarak entegredir. Sözlük bulunmadığında sistem duraksamadan saf BOM ↔ ERP karşılaştırmasını yürütür.

---

## 🏗️ Proje Mimarisi

```text
deka/
├── src/
│   └── bom_enricher/
│       ├── comparison.py       # Çekirdek karşılaştırma, 3 aşamalı eşleme ve Excel raporlama motoru
│       ├── semantic_parser.py  # Elektriksel değer, voltaj ve birim normalizasyon ayrıştırıcısı
│       ├── excel_io.py         # Esnek Excel okuma (başlık satırı tespiti, veri temizleme)
│       ├── models.py           # Veri modelleri ve veri sınıfları (PartRecord, ComparisonResult)
│       ├── web.py              # FastAPI yerel web sunucusu ve indirme uç noktaları
│       └── static/
│           ├── index.html      # Responsive web arayüzü (Deka maskotu entegreli)
│           ├── deka-logo.png   # Deka kurumsal logosu
│           └── deka-bird.png   # İzole edilmiş Deka kuş maskotu
├── data/
│   └── rules.xlsx              # (Opsiyonel) DEKA Malzeme Tipleri ve Özellikleri Sözlüğü
├── examples/                   # Örnek test ve girdi dosyaları
├── reports/                    # Üretilen Master Excel çıktıları
├── tests/                      # Kapsamlı birim ve entegrasyon testleri (61 test)
├── pyproject.toml              # Bağımlılıklar ve proje metaverileri
└── README.md                   # Proje dokümantasyonu
```

---

## 🚀 Hızlı Başlangıç

### 1. Kurulum

Python 3.10 veya üzeri kurulu bir ortamda:

```powershell
# Sanal ortamı oluşturun
py -m venv .venv

# Sanal ortamı etkinleştirin (Windows PowerShell)
.\.venv\Scripts\Activate.ps1

# Bağımlılıkları geliştirici modunda yükleyin
pip install -e ".[dev]"
```

### 2. Web Arayüzünü Başlatma 🌐

Web arayüzünü yerel olarak başlatmak için:

```powershell
.\.venv\Scripts\python.exe -m bom_enricher.web
```

Başlattıktan sonra tarayıcınızda şu adresi açın:  
👉 **`http://127.0.0.1:8000`**

1. **Müşteri BOM** dosyanızı ve **ERP malzeme** listenizi seçin (veya sürükleyip bırakın).
2. **Karşılaştırmayı Başlat** butonuna tıklayın.
3. Sonuçları anlık olarak tarayıcıda inceleyin ve **Excel Raporunu İndir** butonuyla renkli çıktıyı alın.

---

## 💻 Python ile Programatik Kullanım

Motoru kendi scriptleriniz veya veri boru hatlarınız içinde doğrudan çağırabilirsiniz:

```python
from pathlib import Path
from bom_enricher.comparison import compare_workbooks

# Sözlük olmadan saf karşılaştırma
result = compare_workbooks(
    customer_path=Path("data/musteri_bom.xlsx"),
    erp_path=Path("data/erp_listesi.xlsx"),
    output_path=Path("reports/Karsilastirma_Raporu.xlsx")
)

print(f"Toplam İncelenen Parça: {result.total_part_count}")
for row in result.rows:
    if row["Durum"] == "HATA":
        print(f"[!] HATA: {row['Part Number']} -> {row['Hata/Uyarı Durumu']}")
```

---

## 🧪 Testler

Proje genelinde tüm eşleştirme kuralları, parantez genişletmeleri, birim dönüşümleri ve rapor biçimlendirmeleri test kapsamındadır:

```powershell
.\.venv\Scripts\python.exe -m pytest -v
```

```text
collected 61 items

tests/test_cli.py ...                                                    [  4%]
tests/test_comparison.py ....................                            [ 37%]
tests/test_excel_io.py ......                                            [ 47%]
tests/test_nexar_provider.py .............                               [ 68%]
tests/test_output.py ....                                                [ 75%]
tests/test_validation.py ..........                                      [ 91%]
tests/test_web.py .....                                                  [100%]

============================= 61 passed in ~15s =============================
```

---

## 🔒 Güvenlik ve Gizlilik

* Harici bulut servisleri, üçüncü taraf API'ler (OpenAI, Gemini, Nexar vb.) **kullanılmaz**.
* Yüklenen ve karşılaştırılan tüm dosyalar yerel makinenizde kalır.
* DEKA Elektronik Ar-Ge ve kalite standartlarına uygun olarak tasarlanmıştır.

---

<div align="center">
  <sub>Geliştirici: <b>Sudem</b> · DEKA Elektronik Staj Programı © 2026</sub>
</div>

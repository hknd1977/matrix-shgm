# ✈️ Havacılık Eğitimi Lisans Geçiş ve Müfredat Analizi

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![Streamlit](https://img.shields.io/badge/Streamlit-1.30%2B-FF4B4B)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Database](https://img.shields.io/badge/SQLite-3-lightgrey)

Bu proje; Sivil Havacılık Genel Müdürlüğü (SHGM) standartlarındaki havacılık lisans, sertifika ve eğitim matris verilerinin işlenmesi, SQLite veritabanına aktarılması ve öğrenciler ile akademisyenler için analiz edilip görselleştirilmesini sağlayan Python tabanlı bir masaüstü ve web uygulaması ekosistemidir.

---

## 🌟 Öne Çıkan Özellikler

* **📋 Masaüstü Veri Aktarıcı (`gui_matrix_importer.py`):** Panodan (Clipboard) kopyalanan karmaşık havacılık matris verilerini Regex ve dinamik çapa algoritmalarıyla ayrıştırır, `havacilik_egitim.db` SQLite veritabanına kaydeder.
* **🎯 Lisans Geçiş Köprüsü & Emek Analizi:** Mevcut bir lisansa sahip olan öğrencinin en az ek ders yüküyle kazanabileceği hedef lisansları hesaplar (%100 doğrudan kazanılanlar ile kısmi ek ders gerektirenleri ayırır).
* **📚 Müstakil Ders & Not Barajı Süzgeci:** Herhangi bir dersin geçtiği tüm lisansları ve talep edilen min/maks geçme notu barajlarını (≥ X veya ≤ X) filtreler.
* **🕸️ Etkileşimli İlişki Ağı (Network Graph):** Lisanslar ve müfredatlar arasındaki ortak ders yoğunluklarını ve muafiyet bağlarını `NetworkX` ve `Plotly` ile etkileşimli ağ haritasında gösterir (Manuel ve Otomatik N Yakın modları).
* **🔎 Türkçe Uyumlu Akıllı Arama (`smart_search`):** Türkçe karakter uyumsuzluklarını (İ/i, I/ı vb.) tolere ederek kelime öbeklerine göre dinamik sıralama yapar.
* **📤 Özelleştirilmiş Raporlama:** Seçilen lisans ve müfredat analizlerini Türkçe karakter uyumlu **Excel (.xlsx)** ve **PDF (.pdf)** formatlarında anında üretir.

---

## 🛠️ Kurulum ve Çalıştırma

### 1. Depoyu Klonlayın
```bash
git clone [https://github.com/kullanici-adiniz/havacilik-lisans-rehberi.git](https://github.com/kullanici-adiniz/havacilik-lisans-rehberi.git)
cd havacilik-lisans-rehberi

# ✈️ Havacılık Eğitimi Lisans Geçiş ve Müfredat Analizi

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![Streamlit](https://img.shields.io/badge/Streamlit-1.30%2B-FF4B4B)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Database](https://img.shields.io/badge/SQLite-3-lightgrey)

Bu proje; Sivil Havacılık Genel Müdürlüğü (SHGM) standartlarındaki havacılık lisans, sertifika ve eğitim matris verilerinin işlenmesi, SQLite veritabanına aktarılması ve öğrenciler ile akademisyenler için analiz edilip görselleştirilmesini sağlayan Python tabanlı bir masaüstü ve web uygulaması ekosistemidir.

## ⚠️ Yasal Sorumluluk Reddi ve Veri Doğrulama Uyarısı

> **ÖNEMLİ BİLDİRİM VE SORUMLULUK REDDİ (DISCLAIMER):**
> 
> * **Veri Güncelliği ve Hata Payı:** Bu uygulamada yer alan havacılık lisans, ders ve eğitim matris verileri zaman içinde güncelliğini yitirmiş, değişmiş veya aktarım esnasında teknik eksiklikler barındırıyor olabilir. Verilerin anlık güncelliği ve tam doğruluğu garanti edilmemektedir.
> * **Nihai Kontrol Şartı:** Bu sistem sadece bilgilendirme, ön analiz ve akademik danışmanlığa rehberlik etme amacıyla geliştirilmiştir. Tüm işlemlerin, ders muafiyetlerinin ve lisans şartlarının nihai doğrulaması **kesinlikle** T.C. Sivil Havacılık Genel Müdürlüğü (SHGM) resmi bilgi portalı olan [https://kdmorg.shgm.gov.tr/](https://kdmorg.shgm.gov.tr/) adresi üzerinden yapılmalıdır.
> * **Sorumluluk Kabul Edilmez:** Uygulamanın kullanımından, verilerdeki olası eksiklik veya yanlışlıklardan, bu verilere dayanılarak alınan akademik, idari, hukuki veya kişisel kararlardan ve doğabilecek doğrudan/dolaylı zararlardan yazılım geliştiricileri veya proje sahipleri hiçbir şekilde hukuki ve cezai sorumluluk kabul etmemektedir.

---

## 🤖 Yapay Zeka (AI) Kullanımı ve Etik Bildirimi

Bu projenin kod mimarisi, algoritma kurgusu, veri ayrıştırma kalıpları ve kullanıcı arayüzü (UI) bileşenleri **Google Gemini Flash** yapay zeka modeli yardımıyla geliştirilmiştir.

Yapay zeka teknolojisi etik kullanım ilkelerine bağlı kalınarak; kod geliştirme asistanı, algoritma optimizasyonu ve teknik dokümantasyon desteği amacıyla kullanılmıştır.

---

## 🌟 Öne Çıkan Özellikler

* **📋 Masaüstü Veri Aktarıcı (`gui_matrix_importer.py`):** Panodan (Clipboard) kopyalanan karmaşık havacılık matris verilerini Regex ve dinamik çapa algoritmalarıyla ayrıştırır, `havacilik_egitim.db` SQLite veritabanına kaydeder.
* **🎯 Lisans Geçiş Köprüsü & Emek Analizi:** Mevcut bir lisansa sahip olan öğrencinin en az ek ders yüküyle kazanabileceği hedef lisansları hesaplar (%100 doğrudan kazanılanlar ile kısmi ek ders gerektirenleri ayırır).
* **📚 Müstakil Ders & Not Barajı Süzgeci:** Herhangi bir dersin geçtiği tüm lisansları ve talep edilen min/maks geçme notu barajlarını (≥ X veya ≤ X) filtreler.
* **🕸️ Etkileşimli İlişki Ağı (Network Graph):** Lisanslar ve müfredatlar arasındaki ortak ders yoğunluklarını ve muafiyet bağlarını `NetworkX` ve `Plotly` ile etkileşimli ağ haritasında gösterir (Manuel ve Otomatik N Yakın modları).
* **🔎 Türkçe Uyumlu Akıllı Arama (`smart_search`):** Türkçe karakter uyumsuzluklarını (İ/i, I/ı vb.) tolere ederek kelime öbeklerine göre dinamik sıralama yapar.
* **📤 Özelleştirilmiş Raporlama:** Seçilen lisans ve müfredat analizlerini Türkçe karakter uyumlu **Excel (.xlsx)** ve **PDF (.pdf)** formatlarında anında üretir.

---


# MetricusPrime

Performans görevi belgesini (yönerge, konu dağılım listesi, dereceli puanlama anahtarı, değerlendirme çizelgeleri) Word olarak üreten tek sayfalık form.

## Kullanım

1. Sayfayı açın: https://ataturk-mtal.github.io/MetricusPrime/
2. e-Okul > Sınıf Listesi raporunu Excel (XLS) olarak indirip sayfaya sürükleyin. Birden fazla şube eklenebilir.
3. Ders bilgilerini, öğretmenleri ve okul müdürünü yazın.
4. "Word belgesini oluştur" düğmesine basın; belge bilgisayarınıza iner.

Konular ve sorular forma girilirse öğrencilere rastgele dağıtılır ve belgeye yazılır; öğrenci sayısı kadar konu gerekmez. Girilmezse tablolar boş gelir. Görevin amacı ve güvenlik kuralları Word'de doldurulur.

## Gizlilik

Sınıf listesi yalnız tarayıcının içinde okunur; hiçbir yere gönderilmez ve saklanmaz. Tarayıcıda yalnız form alanları (okul, ders, tarihler, öğretmen adları) hatırlanır.

## İnternetsiz kullanım

Depoyu ZIP olarak indirip `index.html` dosyasına çift tıklamak yeterlidir.

## Yapı

- `index.html`: form ve arayüz
- `app.js`: sınıf listesini okuma ve Word belgesini üretme
- `vendor/`: [SheetJS](https://sheetjs.com) (Apache-2.0) ve [docx](https://docx.js.org) (MIT)

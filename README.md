# MetricusPrime — Performans Yönergesi ve Çizelgesi Oluşturma

Performans görevi belgesini (yönerge, konu dağılım listesi, dereceli puanlama anahtarı, değerlendirme çizelgeleri) Word olarak üreten tek sayfalık form.

## Kullanım

1. Sayfayı açın: https://ataturk-mtal.github.io/MetricusPrime/
2. e-Okul > Sınıf Listesi raporunu **"Excel (Sadece Veri)"** biçiminde dışa aktarıp sayfaya sürükleyin. Birden fazla şube eklenebilir.
3. Ders bilgilerini, öğretmenleri ve okul müdürünü yazın.
4. "Word belgesini oluştur" düğmesine basın; belge bilgisayarınıza iner.

Konular ve sorular forma girilirse öğrencilere rastgele dağıtılır ve belgeye yazılır; öğrenci sayısı kadar konu gerekmez. Girilmezse tablolar boş gelir. Görevin amacı ve güvenlik kuralları Word'de doldurulur.

## Örnek rapor

Öğrencilerle paylaşılabilecek 6 sayfalık örnek rapor: https://ataturk-mtal.github.io/MetricusPrime/ornek-rapor.pdf

## Gizlilik ve KVKK

Sayfa, 6698 sayılı KVKK'ya uygun olacak biçimde tasarlanmıştır: hiçbir kişisel veri sunucuya gönderilmez ve sunucuda tutulmaz.

- Sınıf listesi yalnız tarayıcının içinde okunur; öğrenci bilgileri tarayıcıda da saklanmaz.
- Tarayıcıda yalnız form alanları (okul, ders, tarihler, öğretmen adları, konular, tutum maddeleri) hatırlanır; sayfadaki düğmeyle silinebilir.
- Çerez, izleme ya da istatistik aracı yoktur.
- Üretilen Word belgesi öğrenci adlarını içerir; saklanması ve paylaşılması indiren öğretmenin sorumluluğundadır.

## İnternetsiz kullanım

Depoyu ZIP olarak indirip `index.html` dosyasına çift tıklamak yeterlidir.

## Katkı

Sorun açma, dallanma ve birleştirme kuralları: [CONTRIBUTING.md](CONTRIBUTING.md). Geliştirme `dev` dalından ayrılan dallarda yapılır; `dev` üzerinden `main` dalına birleşir.

## Yapı

- `index.html`: form ve arayüz
- `app.js`: sınıf listesini okuma ve Word belgesini üretme
- `vendor/`: [SheetJS](https://sheetjs.com) (Apache-2.0) ve [docx](https://docx.js.org) (MIT)

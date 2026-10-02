# Katkı ve Birleştirme Politikası

Bu belge MetricusPrime deposunda sorun (issue) açma, geliştirme ve birleştirme kurallarını anlatır.

## Dallar

| Dal | Ne için | Kim yazar |
|---|---|---|
| `main` | Yayındaki sürüm. https://ataturk-mtal.github.io/MetricusPrime/ bu daldan yayımlanır. | Yalnız `dev` dalından açılan PR ile güncellenir. |
| `dev` | Bir sonraki sürümün toplandığı dal. Varsayılan dal budur. | Yalnız çalışma dallarından açılan PR ile güncellenir. |
| çalışma dalları | Tek bir iş için açılan kısa ömürlü dallar. | Geliştiren kişi. |

`main` ve `dev` dallarına doğrudan gönderim (push) kapalıdır.

## Akış

1. **Sorun açın.** Her değişiklik bir sorunla başlar (hata ya da öneri). Küçük yazım düzeltmeleri bunun dışındadır.
2. **Dal açın.** Çalışma dalı her zaman güncel `dev` dalından ayrılır:
   ```bash
   git checkout dev && git pull
   git checkout -b fix/12-xls-okuma-hatasi
   ```
3. **Geliştirin ve deneyin.** Aşağıdaki "Birleştirmeden önce" listesini uygulayın.
4. **`dev` dalına PR açın.** PR açıklamasına `Closes #12` yazın; birleşince sorun kendiliğinden kapanır.
5. **`dev` dalına birleştirin.** Yöntem: *Squash and merge*. Çalışma dalı birleşince silinir.
6. **Yayımlayın.** `dev` denenip hazır olduğunda `dev` → `main` PR'ı açılır. Yöntem: *Create a merge commit*.
   Birleşince sayfa birkaç dakika içinde yayına girer.

`main` dalına `dev` dışındaki bir daldan açılan PR, otomatik denetimden geçemez ve birleştirilemez.

## Dal adları

`tür/sorun-no-kısa-açıklama` biçiminde, küçük harf ve tire ile:

- `feat/…` yeni özellik
- `fix/…` hata düzeltme
- `docs/…` belge ve metin
- `chore/…` bakım (kütüphane güncelleme, ayar)

Örnek: `feat/21-konu-sablonlari`, `fix/12-xls-okuma-hatasi`.

## Commit iletileri

```
<tür>: <kısa açıklama>
```

Türler: `feat`, `fix`, `docs`, `refactor`, `test`, `chore`. Açıklama Türkçe yazılır.
Örnek: `fix: sadece veri biçimindeki listede şube bulunamıyordu`.

## Sorun (issue) açma

- **Hata bildirimi:** Ne yaptığınızı, ne beklediğinizi ve ne olduğunu yazın. Tarayıcıyı ve sayfanın altındaki sürüm numarasını belirtin.
- **Öneri:** Hangi ihtiyacı karşılayacağını ve kimin kullanacağını yazın.
- Bir sorun tek bir konuyu anlatır. Aynı sorun daha önce açılmış mı diye bakın.
- **Ekran görüntüsü ya da dosya eklerken öğrenci adlarını ve numaralarını kapatın.** Gerçek sınıf listesi (XLS) ya da üretilmiş belge eklemeyin.

## Birleştirmeden önce

- [ ] Sayfa yerelde açıldı (`index.html` dosyasına çift tıklayarak ya da `python3 -m http.server`) ve tarayıcı konsolunda hata yok.
- [ ] Bir sınıf listesiyle belge üretildi ve Word'de açılıp bakıldı.
- [ ] Depoya öğrenci verisi girmedi: XLS, DOCX, `ogrenciler.json` yok. Örnek ve denemelerde uydurma adlar kullanıldı.
- [ ] Sayfa hiçbir veriyi dışarı göndermiyor: yeni ağ isteği, çerez, izleme ya da dış kaynaklı betik eklenmedi.
- [ ] Kullanıcıya görünen metinler Türkçe ve ders bağımsız.
- [ ] Kullanıcının göreceği bir değişiklikse "Hakkında" bölümündeki sürüm numarası artırıldı.

## Sürüm numarası

`büyük.küçük` biçimindedir (ör. 1.9). Yeni özellik ya da görünür değişiklikte küçük numara artar.
Belgenin yapısını ya da kullanım akışını değiştiren büyük değişikliklerde büyük numara artar.

## Gizlilik kuralı

Bu depo herkese açıktır. Öğrenciye ait hiçbir veri (ad, numara, sınıf listesi, üretilmiş belge) depoya,
sorunlara ya da PR'lara eklenmez. Yanlışlıkla eklenirse hemen alan yetkilisine haber verin; dosyayı silmek yetmez,
geçmişten de temizlenmesi gerekir.

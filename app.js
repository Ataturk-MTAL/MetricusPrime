// Performans görevi belgesi: e-Okul sınıf listesini okur, Word (.docx) belgesini üretir.
// Tarayıcıda window.PerformansGorevi, Node'da module.exports olarak kullanılır.
// Kütüphaneler dışarıdan verilir: XLSX (SheetJS) ve docx (docx.js).
(function (kok, fabrika) {
  if (typeof module === "object" && module.exports) {
    module.exports = fabrika();
  } else {
    kok.PerformansGorevi = fabrika();
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const SINIF_DESENI = /(\d+)\.\s*Sınıf\s*\/\s*(\S+)\s*Şubesi/i;
  const GEREKLI_SUTUNLAR = ["Öğrenci No", "Adı", "Soyadı"];
  const CM = 567; // 1 cm = 567 twip
  const DIKEY_KENAR_CM = 1.8;
  const YATAY_KENAR_CM = 1.4;
  const UST_ALT_KENAR_CM = 1.5;
  const DIKEY_GENISLIK_CM = 21 - 2 * DIKEY_KENAR_CM;
  const YATAY_GENISLIK_CM = 29.7 - 2 * YATAY_KENAR_CM;
  const ETIKET_RENGI = "F2F2F2";
  const YAZI_TIPI = "Times New Roman";
  const EN_AZ_KONU_SATIRI = 10;
  const TUTUM_PUANI = 20; // hazır maddelerin puanı
  const EN_COK_PUAN = 100;

  // (kısa ad, zümre ölçüt metni, puan, [çok iyi, iyi, orta, geliştirilmeli])
  const OLCUTLER = [
    ["Planlama", "Ödevin nasıl yapılacağını planlamıştır", 10, [
      "Çalışma takvimi ve iş adımlarını yazılı hazırlamış, zamanında teslim etmiştir.",
      "Planı hazırlamış, bazı adımlar eksik ya da geç kalmıştır.",
      "Plan yüzeysel, takvim belirsizdir.",
      "Plan hazırlamamıştır."]],
    ["Kaynak çeşitliliği", "Ödevi hazırlarken çeşitli kaynaklardan yararlanmıştır", 10, [
      "Ders kitabı, kitap/makale, kurum sitesi ve güncel örnekler gibi en az 4 farklı kaynak kullanmıştır.",
      "3 farklı kaynak kullanmıştır.",
      "1-2 kaynak kullanmıştır.",
      "Kaynak kullanmamış veya tek kaynaktan kopyalamıştır."]],
    ["Bilgiden yararlanma", "Performans ödevinde çalışılmak üzere seçilen konunun raporlaştırılmasında ulaşılan kaynaklardaki bilgilerden yeterince yararlanmıştır", 10, [
      "Bilgiyi kendi cümleleriyle işlemiş; örnek, veri ve gözlemlerle desteklemiş; soruları gerekçeli cevaplamıştır.",
      "Bilgiyi çoğunlukla işlemiş; soru cevaplarında küçük eksikler vardır.",
      "Bilgi büyük ölçüde aktarılmış, yorum azdır.",
      "Bilgi kopyalanmış veya konuyla ilgisizdir."]],
    ["Görsel materyal", "Performans ödevi çeşitli görsel materyaller ile desteklenmiştir", 15, [
      "Konuya uygun şema/çizim, fotoğraf, tablo ve grafik gibi farklı türde görseller vardır.",
      "En az iki farklı türde görsel vardır.",
      "Yalnız bir tür görsel vardır.",
      "Görsel yoktur veya konuyla ilgisizdir."]],
    ["Görsel açıklaması", "Performans çalışmasında yer alan görsel materyal, konuya ilişkin özet bilgilerle desteklenmiştir", 15, [
      "Her görselin numarası, başlığı ve açıklaması vardır; metinde görsele atıf yapılmıştır.",
      "Görsellerin çoğu açıklanmıştır.",
      "Açıklamalar yetersizdir.",
      "Görseller açıklamasızdır."]],
    ["Kaynakça", "Ödevde kullanılan kaynaklar uygun biçimde rapora yansıtılmıştır", 15, [
      "Kaynakça eksiksiz; yazar/site adı, başlık, bağlantı ve erişim tarihi yazılmıştır.",
      "Kaynakça var, bazı bilgiler eksiktir.",
      "Yalnız bağlantı adresleri yazılmıştır.",
      "Kaynakça yoktur."]],
    ["Yazım ve anlatım", "Rapor anlaşılır biçimde yazılmıştır, oluşturulan cümleler Türkçe yazım kuralına uygundur", 5, [
      "Anlatım akıcı, yazım ve noktalama hatası yok denecek kadar azdır.",
      "Birkaç yazım hatası vardır, anlatım anlaşılırdır.",
      "Yazım hataları anlamayı zorlaştırmaktadır.",
      "Rapor anlaşılmamaktadır."]],
    ["Amaca uygunluk", "Ortaya çıkan performans ödevi, konunun amacına uygundur", 20, [
      "Çalışma / sunum eksiksizdir; sonuçlar yorumlanmış, konu güncel örneklerle ilişkilendirilmiş, soruların tamamı doğru cevaplanmıştır.",
      "Çalışma / sunum yeterlidir; yorum, güncel örnekler veya soru cevaplarından biri kısmen eksiktir.",
      "Çalışma / sunum eksiktir; soruların bir kısmı cevapsızdır.",
      "Çalışma konunun amacını karşılamamaktadır."]],
  ];
  const ARALIK = {
    5: ["5", "4-3", "2", "1-0"],
    10: ["10-9", "8-6", "5-3", "2-0"],
    15: ["15-13", "12-9", "8-5", "4-0"],
    20: ["20-17", "16-12", "11-7", "6-0"],
  };
  const SEVIYELER = ["Çok İyi", "İyi", "Orta", "Geliştirilmeli"];

  const TUTUM_VARSAYILAN = {
    atolye: [
      "Atölye içi tutum ve davranışları; arkadaşlarına ve öğretmenine saygılıdır",
      "Derse etkin katılır, soru sorar, verilen uygulamaları zamanında tamamlar",
      "Ders araç-gereçlerini (defter, kalem, avadanlık) düzenli getirir",
      "İş önlüğü giyer; iş sağlığı ve güvenliği kurallarına uyar, enerji altında çalışmaz",
      "Çalışma masasını, araç-gereci ve atölyeyi düzenli ve temiz bırakır",
    ],
    sinif: [
      "Sınıf içi tutum ve davranışları; arkadaşlarına ve öğretmenine saygılıdır",
      "Derse etkin katılır, soru sorar, verilen çalışmaları zamanında tamamlar",
      "Ders araç-gereçlerini (defter, kalem, ders kitabı) düzenli getirir",
      "Verilen görev ve sorumlulukları zamanında yerine getirir",
      "Sırasını ve dersliği düzenli ve temiz bırakır",
    ],
  };

  const RAPOR_BOLUMLERI = [
    "Kapak (okul, ders, konu, öğrencinin adı-soyadı, sınıfı, numarası, teslim tarihi)",
    "Amaç",
    "Teorik bilgi (kendi cümleleriyle)",
    "Şema, çizim veya görseller; varsa malzeme listesi",
    "Konunun işlenişi: açıklamalar, varsa hesaplamalar, tablolar ve grafikler",
    "Güncel örnekler ve günlük hayattaki / meslekteki kullanım alanları",
    "Soruların cevapları: her soru yazılır, altına gerekçeli cevabı verilir",
    "Sonuç ve yorum (karşılaşılan sorunlar ve çözümleri)",
    "Kaynakça (yazar/site adı, başlık, bağlantı, erişim tarihi)",
  ];

  const GOREV_TURLERI = [
    ["Uygulama + Rapor", "[Uygulamanın nerede ve nasıl yapılacağını, neyin teslim edileceğini yazınız.]"],
    ["Araştırma + Rapor", "Konu kaynaklardan araştırılır; karşılaştırma tabloları ve güncel örneklerle rapor hazırlanır."],
    ["Araştırma + Sunum", "Rapora ek olarak 5-7 dakikalık sunum (8-12 slayt) hazırlanır ve sınıfta sunulur."],
  ];

  const KURALLAR = [
    "Rapor yalnız el yazısıyla hazırlanır; bilgisayar çıktısı rapor kabul edilmez. A4 kâğıt, en az 6 sayfa. Fotoğraf ve benzeri görseller çıktı alınıp yapıştırılabilir; görseller numaralanır ve açıklanır.",
    "Başka bir kaynaktan veya arkadaşından kopyalanan çalışmalar değerlendirmeye alınmaz. Yapay zekâ araçlarından yararlanılabilir; ancak öğrenci raporunu kendi cümleleriyle yazar ve kullandığı aracı kaynakçada belirtir.",
    "Her konu için cevaplanacak sorular verilmiştir. Soruların tamamı cevaplanmadan görev teslim edilmiş sayılmaz.",
    "Çalışma planı ve ara kontrol tarihlerine uyulması, zümre ölçeğinin \"planlama\" ölçütünde değerlendirilir.",
    "Güvenlik: [Derse özgü iş sağlığı ve güvenliği kurallarını yazınız.]",
  ];

  // ── e-Okul sınıf listesi ──────────────────────────────────
  function hucre(deger) {
    return String(deger === undefined || deger === null ? "" : deger).trim();
  }

  /** e-Okul "Sınıf Listesi" (OOG01001R020*.XLS) dosyasını okur. Hata durumunda Türkçe mesajla Error fırlatır. */
  function sinifListesiOku(XLSX, veri, dosyaAdi) {
    let kitap;
    try {
      kitap = XLSX.read(veri, { type: "array" });
    } catch (hata) {
      throw new Error(dosyaAdi + ": dosya açılamadı. e-Okul'dan indirilen Excel (XLS) dosyasını seçtiğinizden emin olun.");
    }
    const sayfa = kitap.Sheets[kitap.SheetNames[0]];
    const satirlar = XLSX.utils.sheet_to_json(sayfa, { header: 1, defval: "", raw: true });

    const baslik = satirlar.slice(0, 6).map((s) => s.map(hucre).join(" ")).join(" ");
    const eslesme = SINIF_DESENI.exec(baslik);
    if (!eslesme) {
      throw new Error(dosyaAdi + ": sınıf/şube bilgisi bulunamadı. Bu dosya e-Okul \"Sınıf Listesi\" raporu olmayabilir.");
    }

    let baslikSatiri = -1;
    let sutun = null;
    for (let r = 0; r < satirlar.length; r++) {
      const degerler = satirlar[r].map(hucre);
      if (GEREKLI_SUTUNLAR.every((ad) => degerler.includes(ad))) {
        baslikSatiri = r;
        sutun = Object.fromEntries(GEREKLI_SUTUNLAR.map((ad) => [ad, degerler.indexOf(ad)]));
        break;
      }
    }
    if (baslikSatiri < 0) {
      throw new Error(dosyaAdi + ": \"Öğrenci No\", \"Adı\", \"Soyadı\" sütunları bulunamadı.");
    }

    const ogrenciler = [];
    for (let r = baslikSatiri + 1; r < satirlar.length; r++) {
      const no = hucre(satirlar[r][sutun["Öğrenci No"]]).replace(/\.0$/, "");
      if (!/^\d+$/.test(no)) continue;
      ogrenciler.push({
        no: Number(no),
        ad: hucre(satirlar[r][sutun["Adı"]]),
        soyad: hucre(satirlar[r][sutun["Soyadı"]]),
      });
    }
    if (ogrenciler.length === 0) {
      throw new Error(dosyaAdi + ": öğrenci satırı bulunamadı.");
    }
    ogrenciler.sort((a, b) => a.no - b.no);
    return { sinif: eslesme[1], sube: eslesme[1] + "/" + eslesme[2], dosya: dosyaAdi, ogrenciler };
  }

  // ── Word yardımcıları ─────────────────────────────────────
  function yardimcilar(docx) {
    const { Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, ShadingType,
      BorderStyle, VerticalAlign } = docx;
    const cm = (deger) => Math.round(deger * CM);
    const KENARSIZ = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
    const KENARSIZ_HUCRE = { top: KENARSIZ, bottom: KENARSIZ, left: KENARSIZ, right: KENARSIZ };

    /** "**kalın**" ve "\n" işaretlerini TextRun dizisine çevirir. */
    function kosular(metin, boyut, kalin) {
      const sonuc = [];
      String(metin).split("**").forEach((parca, i) => {
        parca.split("\n").forEach((satir, j) => {
          if (satir === "" && j === 0) return;
          sonuc.push(new TextRun({ text: satir, bold: kalin || i % 2 === 1, size: boyut * 2, font: YAZI_TIPI, break: j > 0 ? 1 : 0 }));
        });
      });
      return sonuc;
    }

    function paragraf(metin, secenek = {}) {
      const { boyut = 10.5, kalin = false, hiza = AlignmentType.LEFT, sonra = 60, girinti } = secenek;
      return new Paragraph({
        children: kosular(metin, boyut, kalin),
        alignment: hiza,
        spacing: { after: sonra },
        indent: girinti ? { left: cm(girinti), hanging: cm(girinti) } : undefined,
      });
    }

    function hucreYap(metin, genislikCm, secenek = {}) {
      const { boyut = 9, kalin = false, etiket = false, ortala = false, yay = 1, kenarsiz = false } = secenek;
      return new TableCell({
        children: [paragraf(metin, { boyut, kalin: kalin || etiket, hiza: ortala ? AlignmentType.CENTER : AlignmentType.LEFT, sonra: 0 })],
        width: { size: cm(genislikCm), type: WidthType.DXA },
        columnSpan: yay,
        verticalAlign: VerticalAlign.CENTER,
        shading: etiket ? { type: ShadingType.CLEAR, color: "auto", fill: ETIKET_RENGI } : undefined,
        borders: kenarsiz ? KENARSIZ_HUCRE : undefined,
        margins: { top: 30, bottom: 30, left: 70, right: 70 },
      });
    }

    /** Başlık satırı gölgeli, kenarlıklı tablo. `ortala`: ortalanacak sütun dizinleri. */
    function tablo(basliklar, satirlar, genislikler, secenek = {}) {
      const { boyut = 9, ortala = [], ekSatirlar = [] } = secenek;
      const baslikSatiri = new TableRow({
        tableHeader: true,
        children: basliklar.map((b, i) => hucreYap(b, genislikler[i], { boyut, etiket: true, ortala: true })),
      });
      const govde = satirlar.map((satir) => new TableRow({
        children: satir.map((deger, i) => hucreYap(deger, genislikler[i], { boyut, ortala: ortala.includes(i) })),
      }));
      return new Table({
        rows: [baslikSatiri, ...govde, ...ekSatirlar],
        columnWidths: genislikler.map(cm),
        width: { size: cm(genislikler.reduce((a, b) => a + b, 0)), type: WidthType.DXA },
        alignment: AlignmentType.CENTER,
      });
    }

    function kenarsizTablo(hucreler, toplamCm, boyut = 10.5) {
      const genislik = toplamCm / hucreler.length;
      return new Table({
        rows: [new TableRow({ children: hucreler.map((m) => hucreYap(m, genislik, { boyut, ortala: true, kenarsiz: true })) })],
        columnWidths: hucreler.map(() => cm(genislik)),
        width: { size: cm(toplamCm), type: WidthType.DXA },
        alignment: AlignmentType.CENTER,
        borders: { top: KENARSIZ, bottom: KENARSIZ, left: KENARSIZ, right: KENARSIZ, insideHorizontal: KENARSIZ, insideVertical: KENARSIZ },
      });
    }

    function satirYap(hucreler) {
      return new TableRow({ children: hucreler });
    }

    return { cm, paragraf, hucreYap, tablo, kenarsizTablo, satirYap };
  }

  // ── Belge ─────────────────────────────────────────────────
  /**
   * @param docx  docx.js kütüphanesi
   * @param bilgi {okul, ders, egitimYili, donem, verilis, planTeslim, araKontrol, teslim,
   *               ogretmenler: string[], dersTuru: "atolye"|"sinif", tutum: {metin, puan}[],
   *               konular: {baslik, birim, tur, urun, sorular: string[]}[]}
   *              Konu girilmemişse konu, soru ve dağılım tabloları boş satırlarla çıkar.
   *              Her sınıfın `dagilim` dizisi (öğrenci sırası → konu dizini) varsa o kullanılır.
   *              Okul müdürünün adı belgeye yazılmaz; imza yerinde elle yazılır.
   * @param siniflar sinifListesiOku çıktılarının dizisi (şube başına bir tane)
   */
  function belgeOlustur(docx, bilgi, siniflar) {
    const { Document, AlignmentType, PageOrientation } = docx;
    const y = yardimcilar(docx);
    const ORTA = AlignmentType.CENTER;
    const sinifNo = siniflar.length ? siniflar[0].sinif : "…";
    const nokta = (deger) => (deger && String(deger).trim() ? String(deger).trim() : "…………");

    const baslik = (alt) => y.paragraf(
      `${nokta(bilgi.okul).toLocaleUpperCase("tr")}\n${nokta(bilgi.egitimYili)} EĞİTİM-ÖĞRETİM YILI ` +
      `${nokta(bilgi.ders).toLocaleUpperCase("tr")} DERSİ\n${sinifNo}. SINIF ${nokta(bilgi.donem)}. DÖNEM ${alt}`,
      { boyut: 11, kalin: true, hiza: ORTA, sonra: 120 });

    const sayfa = (yatay) => ({
      page: {
        size: { width: y.cm(21), height: y.cm(29.7), orientation: yatay ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT },
        margin: {
          top: y.cm(UST_ALT_KENAR_CM), bottom: y.cm(UST_ALT_KENAR_CM),
          left: y.cm(yatay ? YATAY_KENAR_CM : DIKEY_KENAR_CM), right: y.cm(yatay ? YATAY_KENAR_CM : DIKEY_KENAR_CM),
        },
      },
    });
    const bolumler = [];
    const bolumEkle = (cocuklar, yatay = false) => bolumler.push({ properties: sayfa(yatay), children: cocuklar });

    const konular = konulariDuzenle(bilgi.konular);
    const konuVar = konular.length > 0;
    const dagilimlar = new Map(siniflar.map((sinif) =>
      [sinif, konuVar ? dagilimDuzenle(sinif.dagilim, konular.length, sinif.ogrenciler.length) : null]));
    const konuTekrarli = konuVar && siniflar.some((sinif) => sinif.ogrenciler.length > konular.length);
    const gorevTuruMetni = !konuVar
      ? "Bireysel. Her öğrenciye ayrı konu verilir; konular e-Okul sınıf listesindeki sıraya göre dağıtılır (Konu Dağılım Listesi)."
      : "Bireysel. Konular öğrencilere kura ile (rastgele) dağıtılır (Konu Dağılım Listesi)."
        + (konuTekrarli ? " Aynı konu birden fazla öğrenciye verilebilir; her öğrenci çalışmasını bireysel hazırlar." : "");

    // 1. Yönerge
    const bilgiSatirlari = [
      ["Görevin veriliş tarihi", nokta(bilgi.verilis)],
      ["Çalışma planı teslimi", nokta(bilgi.planTeslim)],
      ["Ara kontrol", nokta(bilgi.araKontrol)],
      ["Görevin teslim tarihi", `**${nokta(bilgi.teslim)}**`],
      ["Görev türü", gorevTuruMetni],
      ["Değerlendirme", "Zümre Performans Değerlendirme Ölçeği (100 puan). Performans notu öğrenciye anında bildirilir."],
    ];
    const bilgiTablosu = new docx.Table({
      rows: bilgiSatirlari.map(([etiket, deger]) => y.satirYap([
        y.hucreYap(etiket, 4.5, { boyut: 10, etiket: true }),
        y.hucreYap(deger, DIKEY_GENISLIK_CM - 4.5, { boyut: 10 }),
      ])),
      columnWidths: [y.cm(4.5), y.cm(DIKEY_GENISLIK_CM - 4.5)],
      width: { size: y.cm(DIKEY_GENISLIK_CM), type: docx.WidthType.DXA },
    });
    const ogretmenler = (bilgi.ogretmenler || []).map((o) => o.trim()).filter(Boolean);
    const imzaAdlari = ogretmenler.length ? ogretmenler : ["[Adı SOYADI]"];
    const IMZA_SUTUN = 3;
    const imzaSatirlari = [];
    for (let i = 0; i < imzaAdlari.length; i += IMZA_SUTUN) {
      imzaSatirlari.push(y.kenarsizTablo(imzaAdlari.slice(i, i + IMZA_SUTUN).map((ad) => `${ad}\nÖğretmen`), DIKEY_GENISLIK_CM));
      imzaSatirlari.push(y.paragraf("", { sonra: 120 }));
    }
    bolumEkle([
      baslik("PERFORMANS GÖREVİ YÖNERGESİ"),
      bilgiTablosu,
      y.paragraf("Görevin Amacı", { boyut: 11, kalin: true, sonra: 60 }),
      y.paragraf("[Bu dönem işlenen öğrenme birimlerini ve görevin öğrenciye ne kazandıracağını 2-3 cümleyle yazınız.]", { hiza: AlignmentType.JUSTIFIED }),
      y.paragraf("Görev Türleri", { boyut: 11, kalin: true }),
      ...GOREV_TURLERI.map(([ad, aciklama]) => y.paragraf(`• **${ad}:** ${aciklama}`, { girinti: 0.4 })),
      y.paragraf("Raporun Bölümleri", { boyut: 11, kalin: true }),
      ...RAPOR_BOLUMLERI.map((b, i) => y.paragraf(`${i + 1}. ${b}`, { girinti: 0.5, sonra: 20 })),
      y.paragraf("Kurallar", { boyut: 11, kalin: true }),
      ...KURALLAR.map((k) => y.paragraf(`• ${k}`, { girinti: 0.4 })),
      y.paragraf("", { sonra: 120 }),
      ...imzaSatirlari,
      y.paragraf(".... / .... / 20....\n**Uygundur**\n\n\n……………………………………\nOkul Müdürü", { hiza: ORTA }),
    ]);

    // 2. Konu listesi ve sorular (formda girilmediyse boş satırlar)
    const enKalabalik = siniflar.reduce((en, s) => Math.max(en, s.ogrenciler.length), 0);
    const konuSatiri = Math.max(EN_AZ_KONU_SATIRI, enKalabalik);
    const numaralar = Array.from({ length: konuSatiri }, (_, i) => i + 1);
    const konuSatirlari = konuVar
      ? konular.map((k, i) => [i + 1, k.baslik, k.birim, k.tur, k.urun])
      : numaralar.map((n) => [n, "", "", "", ""]);
    const soruSatirlari = konuVar
      ? konular.map((k, i) => [i + 1, k.baslik,
        k.sorular.length ? k.sorular.map((soru, j) => `${j + 1}. ${soru}`).join("\n") : "1.\n2."])
      : numaralar.map((n) => [n, "", "1.\n2."]);
    bolumEkle([
      baslik("PERFORMANS GÖREVİ KONULARI"),
      y.tablo(["No", "Konu", "Öğrenme Birimi", "Görev Türü", "Öğrenciden Beklenen Çalışma"],
        konuSatirlari, [1, 5, 3, 2.6, 5.8], { boyut: 8.5, ortala: [0, 3] }),
    ]);
    bolumEkle([
      baslik("PERFORMANS GÖREVİ SORULARI"),
      y.paragraf("Her öğrenci kendi konu numarasındaki soruların **tamamını** raporunun \"Soruların Cevapları\" bölümünde gerekçeli olarak cevaplar.", { boyut: 9 }),
      y.tablo(["No", "Konu", "Cevaplanacak Sorular"],
        soruSatirlari, [1, 5, 11.4], { boyut: 8.5, ortala: [0] }),
    ]);

    // 3. Konu dağılım listesi (şube başına)
    const konuNo = (sinif, sira) => (konuVar ? dagilimlar.get(sinif)[sira] + 1 : sira + 1);
    for (const sinif of siniflar) {
      bolumEkle([
        baslik("PERFORMANS GÖREVİ KONU DAĞILIM LİSTESİ"),
        y.paragraf(`**Sınıf / Şube:** ${sinif.sube}     **Öğrenci sayısı:** ${sinif.ogrenciler.length}     **Teslim:** ${nokta(bilgi.teslim)}`, { boyut: 9 }),
        y.tablo(["S.No", "Öğr. No", "Adı Soyadı", "Konu No", "Konu", "Görev Türü", "Tebellüğ\nİmza"],
          sinif.ogrenciler.map((o, i) => {
            const konu = konuVar ? konular[dagilimlar.get(sinif)[i]] : null;
            return [i + 1, o.no, `${o.ad} ${o.soyad}`, konuNo(sinif, i), konu ? konu.baslik : "", konu ? konu.tur : "", ""];
          }),
          [1, 1.4, 4.4, 1.2, 5, 2.4, 2], { boyut: 8.5, ortala: [0, 1, 3, 5] }),
      ]);
    }

    // 4. Dereceli puanlama anahtarı
    const SEVIYE_CM = 3.05;
    bolumEkle([
      baslik("PERFORMANS GÖREVİ DERECELİ PUANLAMA ANAHTARI"),
      y.paragraf("Ölçütler ve puanları zümrenin Performans Değerlendirme Ölçeği ile aynıdır; seviye betimleri öğrencinin görevden ne beklendiğini önceden bilmesi için eklenmiştir. Bu sayfa görevle birlikte öğrenciye verilir. Betimler derse göre uyarlanabilir.", { boyut: 9.5 }),
      y.tablo(["Ölçüt", "Puan", ...SEVIYELER],
        OLCUTLER.map(([kisa, metin, puan, betimler]) => [
          `**${kisa}**\n${metin}`, `**${puan}**`,
          ...betimler.map((betim, i) => `**(${ARALIK[puan][i]})** ${betim}`),
        ]),
        [4, 1.2, SEVIYE_CM, SEVIYE_CM, SEVIYE_CM, SEVIYE_CM],
        {
          boyut: 8.5, ortala: [1],
          ekSatirlar: [y.satirYap([
            y.hucreYap("TOPLAM: 100", 5.2, { boyut: 8.5, kalin: true, ortala: true, yay: 2 }),
            y.hucreYap("**Not:** Puan aralıkları öğretmene yol göstericidir; her ölçütün puanı ölçüt puanını aşamaz.", 4 * SEVIYE_CM, { boyut: 8.5, yay: 4 }),
          ])],
        }),
    ]);

    // 5. Bireysel değerlendirme formu
    const kimlikTablosu = new docx.Table({
      rows: [
        y.satirYap([y.hucreYap("Adı Soyadı", 3, { boyut: 10, etiket: true }), y.hucreYap("", 5.7, { boyut: 10 }),
          y.hucreYap("Sınıfı / No", 3, { boyut: 10, etiket: true }), y.hucreYap("", 5.7, { boyut: 10 })]),
        y.satirYap([y.hucreYap("Konu No", 3, { boyut: 10, etiket: true }), y.hucreYap("", 5.7, { boyut: 10 }),
          y.hucreYap("Teslim Tarihi", 3, { boyut: 10, etiket: true }), y.hucreYap("", 5.7, { boyut: 10 })]),
        y.satirYap([y.hucreYap("Konu", 3, { boyut: 10, etiket: true }), y.hucreYap("", 14.4, { boyut: 10, yay: 3 })]),
      ],
      columnWidths: [3, 5.7, 3, 5.7].map(y.cm),
      width: { size: y.cm(DIKEY_GENISLIK_CM), type: docx.WidthType.DXA },
    });
    bolumEkle([
      baslik("PERFORMANS DEĞERLENDİRME ÖLÇEĞİ"),
      kimlikTablosu,
      y.paragraf("", { sonra: 120 }),
      y.tablo(["No", "ÖĞRENCİDE GÖZLENECEK ÖZELLİKLER", "PUANI", "ALDIĞI PUAN"],
        OLCUTLER.map(([, metin, puan], i) => [i + 1, metin, puan, ""]), [1, 11.4, 2, 3],
        {
          boyut: 10, ortala: [0, 2, 3],
          ekSatirlar: [y.satirYap([
            y.hucreYap("DEĞERLENDİRME GENEL TOPLAMI", 12.4, { boyut: 10, kalin: true, yay: 2 }),
            y.hucreYap("100", 2, { boyut: 10, kalin: true, ortala: true }),
            y.hucreYap("", 3, { boyut: 10 }),
          ])],
        }),
      y.paragraf("\n**Öğretmen görüşü:**\n\n\n\n"),
      y.kenarsizTablo(["Öğrencinin İmzası", "Ders Öğretmeni\nAdı Soyadı / İmza"], DIKEY_GENISLIK_CM),
    ]);

    // 6. Sınıf çizelgeleri (yatay)
    function topluCizelge(sinif, alt, olcutBasliklari, aciklama, konuSutunu, toplamPuan = EN_COK_PUAN) {
      const sabit = [1, 1.4, 5.4].concat(konuSutunu ? [1.3] : []);
      const TOPLAM_CM = 1.6;
      const olcutCm = (YATAY_GENISLIK_CM - sabit.reduce((a, b) => a + b, 0) - TOPLAM_CM) / olcutBasliklari.length;
      const basliklar = ["S.No", "Öğr. No", "Adı Soyadı"].concat(konuSutunu ? ["Konu No"] : [], olcutBasliklari, [`TOPLAM\n(${toplamPuan})`]);
      const genislikler = sabit.concat(olcutBasliklari.map(() => olcutCm), [TOPLAM_CM]);
      const satirlar = sinif.ogrenciler.map((o, i) => [i + 1, o.no, `${o.ad} ${o.soyad}`]
        .concat(konuSutunu ? [konuNo(sinif, i)] : [], olcutBasliklari.map(() => ""), [""]));
      bolumEkle([
        baslik(alt),
        y.paragraf(`**Sınıf / Şube:** ${sinif.sube}     **Ders Öğretmeni:** ……………………………………     ${aciklama}`, { boyut: 8.5 }),
        y.tablo(basliklar, satirlar, genislikler, { boyut: 8, ortala: basliklar.map((_, i) => i).filter((i) => i !== 2) }),
      ], true);
    }
    for (const sinif of siniflar) {
      topluCizelge(sinif, "PERFORMANS GÖREVİ SINIF DEĞERLENDİRME ÇİZELGESİ",
        OLCUTLER.map(([kisa, , puan]) => `${kisa}\n(${puan})`),
        "Ölçütler zümre Performans Değerlendirme Ölçeği ile aynıdır.", true);
    }

    // 7. Tutum ve davranış ölçeği + çizelgeleri
    const atolye = bilgi.dersTuru !== "sinif";
    const tutum = tutumDuzenle(bilgi.tutum, atolye ? "atolye" : "sinif");
    const tutumToplami = tutumToplam(tutum);
    if (tutumToplami > EN_COK_PUAN) {
      throw new Error(`Tutum ölçeğinin toplam puanı ${tutumToplami}; ${EN_COK_PUAN} puanı geçemez.`);
    }
    bolumEkle([
      baslik(`${atolye ? "ATÖLYE" : "SINIF"} İÇİ TUTUM VE DAVRANIŞ PERFORMANS ÖLÇEĞİ`),
      y.paragraf(`Zümre kararı gereği birinci performans notu; öğrencinin ${atolye ? "atölye" : "sınıf"} içi tutum ve davranışları, derse katılımı ve ders araç-gereçlerini düzenli getirmesi esas alınarak dönem boyunca yapılan gözlemlerle verilir.`, { boyut: 9.5 }),
      y.tablo(["No", "ÖĞRENCİDE GÖZLENECEK ÖZELLİKLER", "PUANI"],
        tutum.map((madde, i) => [i + 1, madde.metin, madde.puan]), [1, 14.4, 2],
        {
          boyut: 10, ortala: [0, 2],
          ekSatirlar: [y.satirYap([
            y.hucreYap("TOPLAM", 15.4, { boyut: 10, kalin: true, yay: 2 }),
            y.hucreYap(String(tutumToplami), 2, { boyut: 10, kalin: true, ortala: true }),
          ])],
        }),
    ]);
    for (const sinif of siniflar) {
      topluCizelge(sinif, "TUTUM VE DAVRANIŞ SINIF DEĞERLENDİRME ÇİZELGESİ",
        tutum.map((madde, i) => `Ölçüt ${i + 1}\n(${madde.puan})`),
        "Ölçüt numaraları bir önceki sayfadaki ölçekle aynıdır.", false, tutumToplami);
    }

    return new Document({
      creator: "Performans Görevi Formu",
      title: `${nokta(bilgi.ders)} Performans Görevi`,
      styles: { default: { document: { run: { font: YAZI_TIPI, size: 21 } } } },
      sections: bolumler,
    });
  }

  /** Başlığı boş konuları atar, alanları kırpar. */
  function konulariDuzenle(konular) {
    const kirp = (deger) => String(deger || "").trim();
    return (konular || [])
      .map((k) => ({
        baslik: kirp(k.baslik), birim: kirp(k.birim), tur: kirp(k.tur), urun: kirp(k.urun),
        sorular: (k.sorular || []).map(kirp).filter(Boolean),
      }))
      .filter((k) => k.baslik);
  }

  /**
   * Konuları öğrencilere rastgele ve dengeli dağıtır; öğrenci sırası → konu dizini döner.
   * Konu azsa her konu eşit sayıda (en çok bir fark) tekrar eder; konu fazlaysa rastgele seçilenler verilir.
   */
  function konuDagit(konuSayisi, ogrenciSayisi, rastgele = Math.random) {
    if (konuSayisi <= 0) return [];
    const karistir = (dizi) => {
      for (let i = dizi.length - 1; i > 0; i--) {
        const j = Math.floor(rastgele() * (i + 1));
        [dizi[i], dizi[j]] = [dizi[j], dizi[i]];
      }
      return dizi;
    };
    const deste = [];
    while (deste.length < ogrenciSayisi) {
      deste.push(...karistir(Array.from({ length: konuSayisi }, (_, i) => i)));
    }
    return karistir(deste.slice(0, ogrenciSayisi));
  }

  function dagilimGecerli(dagilim, konuSayisi, ogrenciSayisi) {
    return Array.isArray(dagilim) && dagilim.length === ogrenciSayisi
      && dagilim.every((d) => Number.isInteger(d) && d >= 0 && d < konuSayisi);
  }

  /** Geçerli bir dağılım yoksa konuları sırayla (tekrar ederek) dağıtır. */
  function dagilimDuzenle(dagilim, konuSayisi, ogrenciSayisi) {
    return dagilimGecerli(dagilim, konuSayisi, ogrenciSayisi)
      ? dagilim : Array.from({ length: ogrenciSayisi }, (_, i) => i % konuSayisi);
  }

  /** Hazır maddeleri {metin, puan} biçiminde verir. */
  function tutumVarsayilan(dersTuru) {
    return TUTUM_VARSAYILAN[dersTuru === "sinif" ? "sinif" : "atolye"].map((metin) => ({ metin, puan: TUTUM_PUANI }));
  }

  /** Boş maddeleri atar, puanı tam sayıya çevirir; hiç madde yoksa hazır maddeleri kullanır. */
  function tutumDuzenle(tutum, dersTuru) {
    const dolu = (tutum || [])
      .map((madde) => ({ metin: String(madde.metin || "").trim(), puan: Math.max(0, Math.round(Number(madde.puan) || 0)) }))
      .filter((madde) => madde.metin);
    return dolu.length ? dolu : tutumVarsayilan(dersTuru);
  }

  function tutumToplam(tutum) {
    return tutum.reduce((toplam, madde) => toplam + (Math.max(0, Math.round(Number(madde.puan) || 0))), 0);
  }

  function dosyaAdi(bilgi, siniflar) {
    const ders = (bilgi.ders || "ders").trim().replace(/[\\/:*?"<>|]+/g, " ").replace(/\s+/g, "-");
    const subeler = siniflar.map((s) => s.sube.replace("/", "")).join("-");
    return `Performans-Gorevi-${ders}${subeler ? "-" + subeler : ""}.docx`;
  }

  return { sinifListesiOku, belgeOlustur, dosyaAdi, tutumVarsayilan, tutumToplam, EN_COK_PUAN,
    konulariDuzenle, konuDagit, dagilimGecerli };
});

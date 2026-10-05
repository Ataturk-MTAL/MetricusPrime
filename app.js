// Performance task document: reads the e-Okul class list and produces the Word (.docx) document.
// Used as window.PerformanceTask in the browser and as module.exports in Node.
// Libraries are passed in from outside: XLSX (SheetJS) and docx (docx.js).
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.PerformanceTask = factory();
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const CLASS_PATTERN = /(\d+)\.\s*Sınıf\s*\/\s*(\S+)\s*Şubesi/i;
  const REQUIRED_COLUMNS = ["Öğrenci No", "Adı", "Soyadı"];
  const DATA_ONLY_HINT = " e-Okul'da raporu \"Excel (Sadece Veri)\" biçiminde dışa aktarıp yeniden deneyin.";
  const TWIPS_PER_CM = 567; // 1 cm = 567 twip
  const PORTRAIT_MARGIN_CM = 1.8;
  const LANDSCAPE_MARGIN_CM = 1.4;
  const TOP_BOTTOM_MARGIN_CM = 1.5;
  const PORTRAIT_WIDTH_CM = 21 - 2 * PORTRAIT_MARGIN_CM;
  const LANDSCAPE_WIDTH_CM = 29.7 - 2 * LANDSCAPE_MARGIN_CM;
  const LABEL_FILL = "F2F2F2";
  const FONT = "Times New Roman";
  const MIN_TOPIC_ROWS = 10;
  const DEFAULT_ATTITUDE_POINTS = 20; // points of each built-in item
  const MAX_POINTS = 100;

  // (short name, department criterion text, points, [very good, good, fair, needs improvement])
  const CRITERIA = [
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
  const SCORE_RANGES = {
    5: ["5", "4-3", "2", "1-0"],
    10: ["10-9", "8-6", "5-3", "2-0"],
    15: ["15-13", "12-9", "8-5", "4-0"],
    20: ["20-17", "16-12", "11-7", "6-0"],
  };
  const LEVELS = ["Çok İyi", "İyi", "Orta", "Geliştirilmeli"];

  const DEFAULT_ATTITUDE_ITEMS = {
    workshop: [
      "Atölye içi tutum ve davranışları; arkadaşlarına ve öğretmenine saygılıdır",
      "Derse etkin katılır, soru sorar, verilen uygulamaları zamanında tamamlar",
      "Ders araç-gereçlerini (defter, kalem, avadanlık) düzenli getirir",
      "İş önlüğü giyer; iş sağlığı ve güvenliği kurallarına uyar, enerji altında çalışmaz",
      "Çalışma masasını, araç-gereci ve atölyeyi düzenli ve temiz bırakır",
    ],
    classroom: [
      "Sınıf içi tutum ve davranışları; arkadaşlarına ve öğretmenine saygılıdır",
      "Derse etkin katılır, soru sorar, verilen çalışmaları zamanında tamamlar",
      "Ders araç-gereçlerini (defter, kalem, ders kitabı) düzenli getirir",
      "Verilen görev ve sorumlulukları zamanında yerine getirir",
      "Sırasını ve dersliği düzenli ve temiz bırakır",
    ],
  };

  const REPORT_SECTIONS = [
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

  const TASK_TYPES = [
    ["Uygulama + Rapor", "[Uygulamanın nerede ve nasıl yapılacağını, neyin teslim edileceğini yazınız.]"],
    ["Araştırma + Rapor", "Konu kaynaklardan araştırılır; karşılaştırma tabloları ve güncel örneklerle rapor hazırlanır."],
    ["Araştırma + Sunum", "Rapora ek olarak 5-7 dakikalık sunum (8-12 slayt) hazırlanır ve sınıfta sunulur."],
    ["Rapor", "Konu kaynaklardan incelenir ve yalnız yazılı rapor hazırlanır; uygulama ya da sunum gerekmez."],
  ];
  const TASK_TYPE_NAMES = TASK_TYPES.map(([name]) => name);
  const DEFAULT_TASK_TYPE = TASK_TYPE_NAMES[0];

  /** Maps a typed task type to its entry in the fixed list (case and whitespace are ignored); null if not found. */
  function matchTaskType(value) {
    const simplify = (text) => String(text || "").toLocaleLowerCase("tr").replace(/\s+/g, "").replace(/ve/g, "+");
    const wanted = simplify(value);
    if (!wanted) return null;
    return TASK_TYPE_NAMES.find((name) => simplify(name) === wanted)
      || (["sadecerapor", "yalnızrapor", "rapor"].includes(wanted) ? "Rapor" : null);
  }

  const RULES = [
    "Rapor yalnız el yazısıyla hazırlanır; bilgisayar çıktısı rapor kabul edilmez. A4 kâğıt, en az 6 sayfa. Fotoğraf ve benzeri görseller çıktı alınıp yapıştırılabilir; görseller numaralanır ve açıklanır.",
    "Başka bir kaynaktan veya arkadaşından kopyalanan çalışmalar değerlendirmeye alınmaz. Yapay zekâ araçlarından yararlanılabilir; ancak öğrenci raporunu kendi cümleleriyle yazar ve kullandığı aracı kaynakçada belirtir.",
    "Her konu için cevaplanacak sorular verilmiştir. Soruların tamamı cevaplanmadan görev teslim edilmiş sayılmaz.",
    "Çalışma planı ve ara kontrol tarihlerine uyulması, zümre ölçeğinin \"planlama\" ölçütünde değerlendirilir.",
    "Güvenlik: [Derse özgü iş sağlığı ve güvenliği kurallarını yazınız.]",
  ];

  // ── e-Okul class list ─────────────────────────────────────
  function cellText(value) {
    return String(value === undefined || value === null ? "" : value).trim();
  }

  /** Reads an e-Okul "Sınıf Listesi" (OOG01001R020*.XLS) file. On failure throws an Error with a Turkish message. */
  function readClassList(XLSX, data, fileName) {
    let workbook;
    try {
      workbook = XLSX.read(data, { type: "array" });
    } catch (error) {
      throw new Error(fileName + ": dosya açılamadı. e-Okul'dan indirilen Excel (XLS) dosyasını seçtiğinizden emin olun.");
    }
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });

    const heading = rows.slice(0, 6).map((s) => s.map(cellText).join(" ")).join(" ");
    const match = CLASS_PATTERN.exec(heading);
    if (!match) {
      throw new Error(fileName + ": sınıf/şube bilgisi bulunamadı. Bu dosya e-Okul \"Sınıf Listesi\" raporu olmayabilir." + DATA_ONLY_HINT);
    }

    let headerRow = -1;
    let column = null;
    for (let r = 0; r < rows.length; r++) {
      const values = rows[r].map(cellText);
      if (REQUIRED_COLUMNS.every((name) => values.includes(name))) {
        headerRow = r;
        column = Object.fromEntries(REQUIRED_COLUMNS.map((name) => [name, values.indexOf(name)]));
        break;
      }
    }
    if (headerRow < 0) {
      throw new Error(fileName + ": \"Öğrenci No\", \"Adı\", \"Soyadı\" sütunları bulunamadı." + DATA_ONLY_HINT);
    }

    const students = [];
    for (let r = headerRow + 1; r < rows.length; r++) {
      const number = cellText(rows[r][column["Öğrenci No"]]).replace(/\.0$/, "");
      if (!/^\d+$/.test(number)) continue;
      students.push({
        number: Number(number),
        firstName: cellText(rows[r][column["Adı"]]),
        lastName: cellText(rows[r][column["Soyadı"]]),
      });
    }
    if (students.length === 0) {
      throw new Error(fileName + ": öğrenci satırı bulunamadı." + DATA_ONLY_HINT);
    }
    students.sort((a, b) => a.number - b.number);
    return { grade: match[1], section: match[1] + "/" + match[2], file: fileName, students };
  }

  // ── Word helpers ──────────────────────────────────────────
  function helpers(docx) {
    const { Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, ShadingType,
      BorderStyle, VerticalAlign } = docx;
    const cm = (value) => Math.round(value * TWIPS_PER_CM);
    const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
    const NO_CELL_BORDERS = { top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER };

    /** Converts "**bold**" and "\n" markers into an array of TextRuns. */
    function runs(text, size, bold) {
      const result = [];
      String(text).split("**").forEach((part, i) => {
        part.split("\n").forEach((line, j) => {
          if (line === "" && j === 0) return;
          result.push(new TextRun({ text: line, bold: bold || i % 2 === 1, size: size * 2, font: FONT, break: j > 0 ? 1 : 0 }));
        });
      });
      return result;
    }

    function paragraph(text, options = {}) {
      const { size = 10.5, bold = false, align = AlignmentType.LEFT, after = 60, indent } = options;
      return new Paragraph({
        children: runs(text, size, bold),
        alignment: align,
        spacing: { after },
        indent: indent ? { left: cm(indent), hanging: cm(indent) } : undefined,
      });
    }

    function makeCell(text, widthCm, options = {}) {
      const { size = 9, bold = false, label = false, center = false, span = 1, borderless = false } = options;
      return new TableCell({
        children: [paragraph(text, { size, bold: bold || label, align: center ? AlignmentType.CENTER : AlignmentType.LEFT, after: 0 })],
        width: { size: cm(widthCm), type: WidthType.DXA },
        columnSpan: span,
        verticalAlign: VerticalAlign.CENTER,
        shading: label ? { type: ShadingType.CLEAR, color: "auto", fill: LABEL_FILL } : undefined,
        borders: borderless ? NO_CELL_BORDERS : undefined,
        margins: { top: 30, bottom: 30, left: 70, right: 70 },
      });
    }

    /** Bordered table with a shaded header row. `center`: indexes of the columns to center. */
    function table(headers, rows, widths, options = {}) {
      const { size = 9, center = [], extraRows = [] } = options;
      const headerRow = new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => makeCell(h, widths[i], { size, label: true, center: true })),
      });
      const body = rows.map((row) => new TableRow({
        children: row.map((value, i) => makeCell(value, widths[i], { size, center: center.includes(i) })),
      }));
      return new Table({
        rows: [headerRow, ...body, ...extraRows],
        columnWidths: widths.map(cm),
        width: { size: cm(widths.reduce((a, b) => a + b, 0)), type: WidthType.DXA },
        alignment: AlignmentType.CENTER,
      });
    }

    function borderlessTable(cells, totalCm, size = 10.5) {
      const width = totalCm / cells.length;
      return new Table({
        rows: [new TableRow({ children: cells.map((m) => makeCell(m, width, { size, center: true, borderless: true })) })],
        columnWidths: cells.map(() => cm(width)),
        width: { size: cm(totalCm), type: WidthType.DXA },
        alignment: AlignmentType.CENTER,
        borders: { top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER, insideHorizontal: NO_BORDER, insideVertical: NO_BORDER },
      });
    }

    function makeRow(cells) {
      return new TableRow({ children: cells });
    }

    return { cm, paragraph, makeCell, table, borderlessTable, makeRow };
  }

  // ── Document ──────────────────────────────────────────────
  /**
   * @param docx  the docx.js library
   * @param form {school, course, academicYear, term, assignedDate, planDueDate, interimCheckDate, dueDate,
   *              teachers: string[], courseType: "workshop"|"classroom", attitude: {text, points}[],
   *              topics: {title, unit, type, deliverable, questions: string[]}[]}
   *              If no topics are entered, the topic, question and distribution tables come out with blank rows.
   *              Each class's `distribution` array (student order → topic index) is used when present.
   *              The principal's name is not written into the document; it is handwritten at the signature.
   * @param classes array of readClassList results (one per section)
   */
  function buildDocument(docx, form, classes) {
    const { Document, AlignmentType, PageOrientation } = docx;
    const h = helpers(docx);
    const CENTER = AlignmentType.CENTER;
    const grade = classes.length ? classes[0].grade : "…";
    const dotted = (value) => (value && String(value).trim() ? String(value).trim() : "…………");

    const heading = (subtitle) => h.paragraph(
      `${dotted(form.school).toLocaleUpperCase("tr")}\n${dotted(form.academicYear)} EĞİTİM-ÖĞRETİM YILI ` +
      `${dotted(form.course).toLocaleUpperCase("tr")} DERSİ\n${grade}. SINIF ${dotted(form.term)}. DÖNEM ${subtitle}`,
      { size: 11, bold: true, align: CENTER, after: 120 });

    const page = (landscape) => ({
      page: {
        size: { width: h.cm(21), height: h.cm(29.7), orientation: landscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT },
        margin: {
          top: h.cm(TOP_BOTTOM_MARGIN_CM), bottom: h.cm(TOP_BOTTOM_MARGIN_CM),
          left: h.cm(landscape ? LANDSCAPE_MARGIN_CM : PORTRAIT_MARGIN_CM), right: h.cm(landscape ? LANDSCAPE_MARGIN_CM : PORTRAIT_MARGIN_CM),
        },
      },
    });
    const sections = [];
    const addSection = (children, landscape = false) => sections.push({ properties: page(landscape), children });

    const topics = normalizeTopics(form.topics);
    const hasTopics = topics.length > 0;
    const distributions = new Map(classes.map((cls) =>
      [cls, hasTopics ? normalizeDistribution(cls.distribution, topics.length, cls.students.length) : null]));
    const topicsRepeat = hasTopics && classes.some((cls) => cls.students.length > topics.length);
    const taskTypeText = !hasTopics
      ? "Bireysel. Her öğrenciye ayrı konu verilir; konular e-Okul sınıf listesindeki sıraya göre dağıtılır (Konu Dağılım Listesi)."
      : "Bireysel. Konular öğrencilere kura ile (rastgele) dağıtılır (Konu Dağılım Listesi)."
        + (topicsRepeat ? " Aynı konu birden fazla öğrenciye verilebilir; her öğrenci çalışmasını bireysel hazırlar." : "");

    // 1. Instructions
    const infoRows = [
      ["Görevin veriliş tarihi", dotted(form.assignedDate)],
      ["Çalışma planı teslimi", dotted(form.planDueDate)],
      ["Ara kontrol", dotted(form.interimCheckDate)],
      ["Görevin teslim tarihi", `**${dotted(form.dueDate)}**`],
      ["Görev türü", taskTypeText],
      ["Değerlendirme", "Zümre Performans Değerlendirme Ölçeği (100 puan). Performans notu öğrenciye anında bildirilir."],
    ];
    const infoTable = new docx.Table({
      rows: infoRows.map(([label, value]) => h.makeRow([
        h.makeCell(label, 4.5, { size: 10, label: true }),
        h.makeCell(value, PORTRAIT_WIDTH_CM - 4.5, { size: 10 }),
      ])),
      columnWidths: [h.cm(4.5), h.cm(PORTRAIT_WIDTH_CM - 4.5)],
      width: { size: h.cm(PORTRAIT_WIDTH_CM), type: docx.WidthType.DXA },
    });
    const teachers = (form.teachers || []).map((t) => t.trim()).filter(Boolean);
    const signatureNames = teachers.length ? teachers : ["[Adı SOYADI]"];
    const SIGNATURE_COLUMNS = 3;
    const signatureRows = [];
    for (let i = 0; i < signatureNames.length; i += SIGNATURE_COLUMNS) {
      signatureRows.push(h.borderlessTable(signatureNames.slice(i, i + SIGNATURE_COLUMNS).map((name) => `${name}\nÖğretmen`), PORTRAIT_WIDTH_CM));
      signatureRows.push(h.paragraph("", { after: 120 }));
    }
    addSection([
      heading("PERFORMANS GÖREVİ YÖNERGESİ"),
      infoTable,
      h.paragraph("Görevin Amacı", { size: 11, bold: true, after: 60 }),
      h.paragraph("[Bu dönem işlenen öğrenme birimlerini ve görevin öğrenciye ne kazandıracağını 2-3 cümleyle yazınız.]", { align: AlignmentType.JUSTIFIED }),
      h.paragraph("Görev Türleri", { size: 11, bold: true }),
      ...TASK_TYPES.filter(([name]) => !hasTopics || topics.some((t) => t.type === name))
        .map(([name, description]) => h.paragraph(`• **${name}:** ${description}`, { indent: 0.4 })),
      h.paragraph("Raporun Bölümleri", { size: 11, bold: true }),
      ...REPORT_SECTIONS.map((s, i) => h.paragraph(`${i + 1}. ${s}`, { indent: 0.5, after: 20 })),
      h.paragraph("Kurallar", { size: 11, bold: true }),
      ...RULES.map((rule) => h.paragraph(`• ${rule}`, { indent: 0.4 })),
      h.paragraph("", { after: 120 }),
      ...signatureRows,
      h.paragraph(".... / .... / 20....\n**Uygundur**\n\n\n……………………………………\nOkul Müdürü", { align: CENTER }),
    ]);

    // 2. Topic list and questions (blank rows if not entered in the form)
    const largestClass = classes.reduce((max, c) => Math.max(max, c.students.length), 0);
    const topicRowCount = Math.max(MIN_TOPIC_ROWS, largestClass);
    const numbers = Array.from({ length: topicRowCount }, (_, i) => i + 1);
    const topicRows = hasTopics
      ? topics.map((t, i) => [i + 1, t.title, t.unit, t.type, t.deliverable])
      : numbers.map((n) => [n, "", "", "", ""]);
    const questionRows = hasTopics
      ? topics.map((t, i) => [i + 1, t.title,
        t.questions.length ? t.questions.map((question, j) => `${j + 1}. ${question}`).join("\n") : "1.\n2."])
      : numbers.map((n) => [n, "", "1.\n2."]);
    addSection([
      heading("PERFORMANS GÖREVİ KONULARI"),
      h.table(["No", "Konu", "Öğrenme Birimi", "Görev Türü", "Öğrenciden Beklenen Çalışma"],
        topicRows, [1, 5, 3, 2.6, 5.8], { size: 8.5, center: [0, 3] }),
    ]);
    addSection([
      heading("PERFORMANS GÖREVİ SORULARI"),
      h.paragraph("Her öğrenci kendi konu numarasındaki soruların **tamamını** raporunun \"Soruların Cevapları\" bölümünde gerekçeli olarak cevaplar.", { size: 9 }),
      h.table(["No", "Konu", "Cevaplanacak Sorular"],
        questionRows, [1, 5, 11.4], { size: 8.5, center: [0] }),
    ]);

    // 3. Topic distribution list (one per section)
    const topicNumber = (cls, index) => (hasTopics ? distributions.get(cls)[index] + 1 : index + 1);
    for (const cls of classes) {
      addSection([
        heading("PERFORMANS GÖREVİ KONU DAĞILIM LİSTESİ"),
        h.paragraph(`**Sınıf / Şube:** ${cls.section}     **Öğrenci sayısı:** ${cls.students.length}     **Teslim:** ${dotted(form.dueDate)}`, { size: 9 }),
        h.table(["S.No", "Öğr. No", "Adı Soyadı", "Konu No", "Konu", "Görev Türü", "Tebellüğ\nİmza"],
          cls.students.map((s, i) => {
            const topic = hasTopics ? topics[distributions.get(cls)[i]] : null;
            return [i + 1, s.number, `${s.firstName} ${s.lastName}`, topicNumber(cls, i), topic ? topic.title : "", topic ? topic.type : "", ""];
          }),
          [1, 1.4, 4.4, 1.2, 5, 2.4, 2], { size: 8.5, center: [0, 1, 3, 5] }),
      ]);
    }

    // 4. Graded scoring rubric
    const LEVEL_CM = 3.05;
    addSection([
      heading("PERFORMANS GÖREVİ DERECELİ PUANLAMA ANAHTARI"),
      h.paragraph("Ölçütler ve puanları zümrenin Performans Değerlendirme Ölçeği ile aynıdır; seviye betimleri öğrencinin görevden ne beklendiğini önceden bilmesi için eklenmiştir. Bu sayfa görevle birlikte öğrenciye verilir. Betimler derse göre uyarlanabilir.", { size: 9.5 }),
      h.table(["Ölçüt", "Puan", ...LEVELS],
        CRITERIA.map(([shortName, text, points, descriptors]) => [
          `**${shortName}**\n${text}`, `**${points}**`,
          ...descriptors.map((descriptor, i) => `**(${SCORE_RANGES[points][i]})** ${descriptor}`),
        ]),
        [4, 1.2, LEVEL_CM, LEVEL_CM, LEVEL_CM, LEVEL_CM],
        {
          size: 8.5, center: [1],
          extraRows: [h.makeRow([
            h.makeCell("TOPLAM: 100", 5.2, { size: 8.5, bold: true, center: true, span: 2 }),
            h.makeCell("**Not:** Puan aralıkları öğretmene yol göstericidir; her ölçütün puanı ölçüt puanını aşamaz.", 4 * LEVEL_CM, { size: 8.5, span: 4 }),
          ])],
        }),
    ]);

    // 5. Individual assessment form
    const identityTable = new docx.Table({
      rows: [
        h.makeRow([h.makeCell("Adı Soyadı", 3, { size: 10, label: true }), h.makeCell("", 5.7, { size: 10 }),
          h.makeCell("Sınıfı / No", 3, { size: 10, label: true }), h.makeCell("", 5.7, { size: 10 })]),
        h.makeRow([h.makeCell("Konu No", 3, { size: 10, label: true }), h.makeCell("", 5.7, { size: 10 }),
          h.makeCell("Teslim Tarihi", 3, { size: 10, label: true }), h.makeCell("", 5.7, { size: 10 })]),
        h.makeRow([h.makeCell("Konu", 3, { size: 10, label: true }), h.makeCell("", 14.4, { size: 10, span: 3 })]),
      ],
      columnWidths: [3, 5.7, 3, 5.7].map(h.cm),
      width: { size: h.cm(PORTRAIT_WIDTH_CM), type: docx.WidthType.DXA },
    });
    addSection([
      heading("PERFORMANS DEĞERLENDİRME ÖLÇEĞİ"),
      identityTable,
      h.paragraph("", { after: 120 }),
      h.table(["No", "ÖĞRENCİDE GÖZLENECEK ÖZELLİKLER", "PUANI", "ALDIĞI PUAN"],
        CRITERIA.map(([, text, points], i) => [i + 1, text, points, ""]), [1, 11.4, 2, 3],
        {
          size: 10, center: [0, 2, 3],
          extraRows: [h.makeRow([
            h.makeCell("DEĞERLENDİRME GENEL TOPLAMI", 12.4, { size: 10, bold: true, span: 2 }),
            h.makeCell("100", 2, { size: 10, bold: true, center: true }),
            h.makeCell("", 3, { size: 10 }),
          ])],
        }),
      h.paragraph("\n**Öğretmen görüşü:**\n\n\n\n"),
      h.borderlessTable(["Öğrencinin İmzası", "Ders Öğretmeni\nAdı Soyadı / İmza"], PORTRAIT_WIDTH_CM),
    ]);

    // 6. Class score sheets (landscape)
    function classScoreSheet(cls, subtitle, criterionHeaders, note, hasTopicColumn, totalPoints = MAX_POINTS) {
      const fixed = [1, 1.4, 5.4].concat(hasTopicColumn ? [1.3] : []);
      const TOTAL_CM = 1.6;
      const criterionCm = (LANDSCAPE_WIDTH_CM - fixed.reduce((a, b) => a + b, 0) - TOTAL_CM) / criterionHeaders.length;
      const headers = ["S.No", "Öğr. No", "Adı Soyadı"].concat(hasTopicColumn ? ["Konu No"] : [], criterionHeaders, [`TOPLAM\n(${totalPoints})`]);
      const widths = fixed.concat(criterionHeaders.map(() => criterionCm), [TOTAL_CM]);
      const rows = cls.students.map((s, i) => [i + 1, s.number, `${s.firstName} ${s.lastName}`]
        .concat(hasTopicColumn ? [topicNumber(cls, i)] : [], criterionHeaders.map(() => ""), [""]));
      addSection([
        heading(subtitle),
        h.paragraph(`**Sınıf / Şube:** ${cls.section}     **Ders Öğretmeni:** ……………………………………     ${note}`, { size: 8.5 }),
        h.table(headers, rows, widths, { size: 8, center: headers.map((_, i) => i).filter((i) => i !== 2) }),
      ], true);
    }
    for (const cls of classes) {
      classScoreSheet(cls, "PERFORMANS GÖREVİ SINIF DEĞERLENDİRME ÇİZELGESİ",
        CRITERIA.map(([shortName, , points]) => `${shortName}\n(${points})`),
        "Ölçütler zümre Performans Değerlendirme Ölçeği ile aynıdır.", true);
    }

    // 7. Attitude and behaviour scale + score sheets
    const isWorkshop = form.courseType !== "classroom";
    const attitude = normalizeAttitude(form.attitude, isWorkshop ? "workshop" : "classroom");
    const attitudeSum = attitudeTotal(attitude);
    if (attitudeSum > MAX_POINTS) {
      throw new Error(`Tutum ölçeğinin toplam puanı ${attitudeSum}; ${MAX_POINTS} puanı geçemez.`);
    }
    addSection([
      heading(`${isWorkshop ? "ATÖLYE" : "SINIF"} İÇİ TUTUM VE DAVRANIŞ PERFORMANS ÖLÇEĞİ`),
      h.paragraph(`Zümre kararı gereği birinci performans notu; öğrencinin ${isWorkshop ? "atölye" : "sınıf"} içi tutum ve davranışları, derse katılımı ve ders araç-gereçlerini düzenli getirmesi esas alınarak dönem boyunca yapılan gözlemlerle verilir.`, { size: 9.5 }),
      h.table(["No", "ÖĞRENCİDE GÖZLENECEK ÖZELLİKLER", "PUANI"],
        attitude.map((item, i) => [i + 1, item.text, item.points]), [1, 14.4, 2],
        {
          size: 10, center: [0, 2],
          extraRows: [h.makeRow([
            h.makeCell("TOPLAM", 15.4, { size: 10, bold: true, span: 2 }),
            h.makeCell(String(attitudeSum), 2, { size: 10, bold: true, center: true }),
          ])],
        }),
    ]);
    for (const cls of classes) {
      classScoreSheet(cls, "TUTUM VE DAVRANIŞ SINIF DEĞERLENDİRME ÇİZELGESİ",
        attitude.map((item, i) => `Ölçüt ${i + 1}\n(${item.points})`),
        "Ölçüt numaraları bir önceki sayfadaki ölçekle aynıdır.", false, attitudeSum);
    }

    return new Document({
      creator: "Performans Görevi Formu",
      title: `${dotted(form.course)} Performans Görevi`,
      styles: { default: { document: { run: { font: FONT, size: 21 } } } },
      sections,
    });
  }

  /** Drops topics with an empty title and trims the fields. */
  function normalizeTopics(topics) {
    const trim = (value) => String(value || "").trim();
    return (topics || [])
      .map((t) => ({
        title: trim(t.title), unit: trim(t.unit), type: matchTaskType(t.type) || DEFAULT_TASK_TYPE, deliverable: trim(t.deliverable),
        questions: (t.questions || []).map(trim).filter(Boolean),
      }))
      .filter((t) => t.title);
  }

  /**
   * Distributes topics to students randomly and evenly; returns student order → topic index.
   * With fewer topics, each topic repeats equally often (at most one apart); with more topics, a random subset is used.
   */
  function distributeTopics(topicCount, studentCount, random = Math.random) {
    if (topicCount <= 0) return [];
    const shuffle = (array) => {
      for (let i = array.length - 1; i > 0; i--) {
        const j = Math.floor(random() * (i + 1));
        [array[i], array[j]] = [array[j], array[i]];
      }
      return array;
    };
    const deck = [];
    while (deck.length < studentCount) {
      deck.push(...shuffle(Array.from({ length: topicCount }, (_, i) => i)));
    }
    return shuffle(deck.slice(0, studentCount));
  }

  function isValidDistribution(distribution, topicCount, studentCount) {
    return Array.isArray(distribution) && distribution.length === studentCount
      && distribution.every((d) => Number.isInteger(d) && d >= 0 && d < topicCount);
  }

  /** Without a valid distribution, hands out topics in order (repeating). */
  function normalizeDistribution(distribution, topicCount, studentCount) {
    return isValidDistribution(distribution, topicCount, studentCount)
      ? distribution : Array.from({ length: studentCount }, (_, i) => i % topicCount);
  }

  /** Returns the built-in items as {text, points}. */
  function defaultAttitude(courseType) {
    return DEFAULT_ATTITUDE_ITEMS[courseType === "classroom" ? "classroom" : "workshop"].map((text) => ({ text, points: DEFAULT_ATTITUDE_POINTS }));
  }

  /** Drops empty items and rounds points to integers; falls back to the built-in items when none remain. */
  function normalizeAttitude(attitude, courseType) {
    const filled = (attitude || [])
      .map((item) => ({ text: String(item.text || "").trim(), points: Math.max(0, Math.round(Number(item.points) || 0)) }))
      .filter((item) => item.text);
    return filled.length ? filled : defaultAttitude(courseType);
  }

  function attitudeTotal(attitude) {
    return attitude.reduce((total, item) => total + (Math.max(0, Math.round(Number(item.points) || 0))), 0);
  }

  function documentFileName(form, classes) {
    const course = (form.course || "ders").trim().replace(/[\\/:*?"<>|]+/g, " ").replace(/\s+/g, "-");
    const sections = classes.map((c) => c.section.replace("/", "")).join("-");
    return `Performans-Gorevi-${course}${sections ? "-" + sections : ""}.docx`;
  }

  return { readClassList, buildDocument, documentFileName, defaultAttitude, attitudeTotal, MAX_POINTS,
    normalizeTopics, distributeTopics, isValidDistribution, TASK_TYPE_NAMES, matchTaskType };
});

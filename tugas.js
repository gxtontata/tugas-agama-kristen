const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  HeadingLevel, PageBreak, LevelFormat, BorderStyle,
  PageNumber, NumberFormat, Header, Footer, TabStopType, TabStopPosition,
  UnderlineType, TableOfContents
} = require('docx');
const fs = require('fs');

const center = (children, spacing) =>
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: spacing || { after: 120 }, children });
 
const justify = (children, spacing) =>
  new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: spacing || { before: 60, after: 120, line: 360, lineRule: 'auto' }, children });
 
const h1 = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    alignment: AlignmentType.CENTER,
    spacing: { before: 240, after: 200 },
    children: [new TextRun({ text, bold: true, size: 28, font: 'Times New Roman', allCaps: true })]
  });
 
const h2 = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 120 },
    children: [new TextRun({ text, bold: true, size: 24, font: 'Times New Roman' })]
  });
 
const h3 = (text) =>
  new Paragraph({
    spacing: { before: 160, after: 100 },
    children: [new TextRun({ text, bold: true, italics: true, size: 24, font: 'Times New Roman' })]
  });
 
const body = (text, extra) =>
  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    indent: { firstLine: 720 },
    spacing: { line: 360, lineRule: 'auto', after: 120 },
    children: [new TextRun({ text, size: 24, font: 'Times New Roman', ...extra })]
  });
 
const bodyNoIndent = (text, extra) =>
  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 360, lineRule: 'auto', after: 120 },
    children: [new TextRun({ text, size: 24, font: 'Times New Roman', ...extra })]
  });
 
const empty = (n = 1) => Array(n).fill(0).map(() =>
  new Paragraph({ children: [new TextRun({ text: '' })] })
);
 
const pageBreak = () =>
  new Paragraph({ children: [new PageBreak()] });

const doc = new Document({
  features: {
    updateFields: true,
  },
  styles: {
    default: {
      document: { run: { font: 'Times New Roman', size: 24 } }
    },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 28, bold: true, font: 'Times New Roman' },
        paragraph: { spacing: { before: 240, after: 200 }, outlineLevel: 0 }
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 24, bold: true, font: 'Times New Roman' },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 1 }
      },
      {
        id: 'TOC1', name: 'TOC Level 1', basedOn: 'Normal', quickFormat: true,
        run: { size: 24, font: 'Times New Roman', bold: true },
        paragraph: { spacing: { after: 80 }, outlineLevel: 0 }
      },
      {
        id: 'TOC2', name: 'TOC Level 2', basedOn: 'Normal', quickFormat: true,
        run: { size: 24, font: 'Times New Roman' },
        paragraph: { spacing: { after: 80 }, outlineLevel: 1 }
      }
    ]
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1701, right: 1134, bottom: 1134, left: 1701 }
        }
      },
      children: [
        ...empty(3),
        center([new TextRun({ text: 'MAKALAH', bold: true, size: 36, font: 'Times New Roman', allCaps: true })], { after: 120 }),
        ...empty(1),
        center([new TextRun({ text: 'TANGGUNG JAWAB UMAT KRISTEN', bold: true, size: 32, font: 'Times New Roman' })], { after: 80 }),
        center([new TextRun({ text: 'DALAM MENJAGA ALAM', bold: true, size: 32, font: 'Times New Roman' })], { after: 200 }),
        ...empty(1),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: '2E8B57', space: 1 } },
          children: [new TextRun({ text: '', size: 24 })]
        }),
        ...empty(4),
        center([new TextRun({ text: 'Disusun untuk Memenuhi Tugas Mata Pelajaran', size: 24, font: 'Times New Roman', italics: true })]),
        center([new TextRun({ text: 'Pendidikan Agama Kristen', size: 24, font: 'Times New Roman', italics: true })], { after: 240 }),
        ...empty(4),
        center([new TextRun({ text: 'Tahun Pelajaran 2025/2026', bold: true, size: 24, font: 'Times New Roman' })], { after: 80 }),
        ...empty(6),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: '2E8B57', space: 1 } },
          children: [new TextRun({ text: '', size: 24 })]
        }),
      ]
    },
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1701, right: 1134, bottom: 1134, left: 1701 }
        }
      },
      children: [
        h1('DAFTAR ISI'),
        new Paragraph({
          spacing: { after: 200 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: '2E8B57', space: 4 } },
          children: []
        }),
        ...empty(2),
        new TableOfContents("Daftar Isi", {
          hyperlink: true,
          headingStyleRange: "1-3",
          stylesWithLevels: [
            { style: "TOC1", level: 1 },
            { style: "TOC2", level: 2 },
          ],
        }),
        
        ...empty(2),
        pageBreak(),
      ]
    },
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1701, right: 1134, bottom: 1134, left: 1701 }
        }
      },
      children: [
        h1('BAB I'),
        h1('PENDAHULUAN'),
 
        h2('A. Latar Belakang'),
        body('Alam semesta adalah karya agung Allah yang mencerminkan kemuliaan dan kebijaksanaan-Nya. Langit, laut, pegunungan, hutan, dan seluruh makhluk hidup di dalamnya merupakan bagian dari ciptaan yang dinyatakan Allah sebagai "sungguh amat baik" (Kejadian 1:31). Namun, dalam perjalanan sejarah umat manusia, alam yang indah ini mengalami kerusakan yang semakin mengkhawatirkan akibat ulah tangan manusia sendiri.'),
        body('Perubahan iklim, deforestasi, pencemaran air dan udara, kepunahan spesies, serta pengelolaan sampah yang buruk adalah sejumlah persoalan lingkungan yang tengah mengancam keberlangsungan hidup di bumi ini. Indonesia, sebagai negara dengan keanekaragaman hayati terbesar kedua di dunia, turut menghadapi tantangan serius dalam upaya pelestarian alam. Pembukaan lahan secara masif, penambangan ilegal, dan polusi industri telah mengakibatkan kerusakan ekosistem yang signifikan.'),
        body('Di tengah krisis lingkungan yang kian mengkhawatirkan tersebut, umat Kristen dipanggil untuk tidak berdiam diri. Iman Kristen memandang alam bukan sekadar sumber daya yang dapat dieksploitasi sesuka hati, melainkan sebagai titipan Allah yang harus dijaga, dipelihara, dan dikelola dengan penuh tanggung jawab. Alkitab sendiri memberikan dasar yang kuat bagi tanggung jawab manusia—termasuk umat Kristen—dalam memelihara ciptaan Allah.'),
        body('Makalah ini disusun untuk menggali lebih dalam bagaimana pandangan iman Kristen terhadap alam semesta, serta apa saja bentuk tanggung jawab konkret yang harus dilakukan umat Kristen dalam menjaga dan melestarikan alam. Dengan memahami hal ini, diharapkan setiap orang percaya dapat menjalankan peran sebagai "penjaga alam" yang sejati, sebagaimana yang dikehendaki oleh Sang Pencipta.'),
 
        h2('B. Rumusan Masalah'),
        bodyNoIndent('Berdasarkan latar belakang yang telah dipaparkan, rumusan masalah dalam makalah ini adalah:'),
        new Paragraph({
          numbering: { reference: 'numbers', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 80 },
          children: [new TextRun({ text: 'Bagaimana pandangan Alkitab mengenai alam sebagai ciptaan Allah?', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          numbering: { reference: 'numbers', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 80 },
          children: [new TextRun({ text: 'Apa saja bentuk tanggung jawab umat Kristen dalam menjaga alam?', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          numbering: { reference: 'numbers', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 200 },
          children: [new TextRun({ text: 'Bagaimana wujud nyata umat Kristen dalam melestarikan lingkungan hidup?', size: 24, font: 'Times New Roman' })]
        }),
 
        h2('C. Tujuan Penulisan'),
        bodyNoIndent('Makalah ini bertujuan untuk:'),
        new Paragraph({
          numbering: { reference: 'numbers2', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 80 },
          children: [new TextRun({ text: 'Menjelaskan pandangan Alkitab mengenai alam sebagai ciptaan Allah yang harus dijaga.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          numbering: { reference: 'numbers2', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 80 },
          children: [new TextRun({ text: 'Menguraikan tanggung jawab umat Kristen dalam melestarikan alam berdasarkan firman Tuhan.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          numbering: { reference: 'numbers2', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 80 },
          children: [new TextRun({ text: 'Memberikan gambaran konkret tindakan nyata yang dapat dilakukan umat Kristen dalam menjaga lingkungan.', size: 24, font: 'Times New Roman' })]
        }),
      ]
    },
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1701, right: 1134, bottom: 1134, left: 1701 }
        }
      },
      children: [
        h1('BAB II'),
        h1('PEMBAHASAN'),
 
        h2('A. Alam sebagai Ciptaan Allah'),
        body('Alkitab secara jelas menyatakan bahwa seluruh alam semesta adalah ciptaan Allah. Dalam kitab Kejadian pasal 1, dikisahkan bagaimana Allah menciptakan langit dan bumi, terang dan gelap, air dan daratan, tumbuhan, bintang-bintang, hewan-hewan, dan akhirnya manusia. Pada setiap tahapan penciptaan tersebut, Allah menyebutkan bahwa ciptaan-Nya itu "baik" (tob), dan setelah seluruh ciptaan selesai, Allah menyatakan semuanya "sungguh amat baik" (Kej. 1:31).'),
        body('Mazmur 19:1-2 mengungkapkan: "Langit menceritakan kemuliaan Allah, dan cakrawala memberitakan pekerjaan tangan-Nya." Alam bukan sekadar materi yang tidak memiliki nilai, melainkan menjadi wahyu umum yang menyatakan keberadaan, kemuliaan, dan kebijaksanaan Allah kepada seluruh umat manusia. Kolose 1:16-17 menegaskan bahwa segala sesuatu diciptakan oleh dan untuk Kristus, dan di dalam Dia segala sesuatu ada dalam keadaan terpadu.'),
        body('Teologi penciptaan dalam iman Kristen memandang alam bukan sebagai milik manusia, melainkan sebagai milik Allah (Maz. 24:1 — "Tuhanlah yang empunya bumi serta segala isinya"). Manusia hanya diberikan kepercayaan untuk mengelolanya. Pandangan ini secara fundamental membedakan iman Kristen dari pandangan materialistis yang melihat alam hanya sebagai objek eksploitasi demi keuntungan ekonomi semata.'),
        body('Dengan memahami alam sebagai ciptaan Allah yang baik dan berharga, umat Kristen dipanggil untuk memperlakukan alam dengan rasa hormat, bukan sebagai benda mati yang bisa diperlakukan sesuka hati. Kerusakan alam, dalam perspektif ini, merupakan bentuk penghinaan terhadap sang Pencipta sendiri.'),
 
        h2('B. Mandat Allah kepada Manusia untuk Memelihara Alam'),
        body('Allah memberikan mandat yang jelas kepada manusia sejak awal penciptaan. Dalam Kejadian 1:28, Allah berfirman: "Beranakcuculah dan bertambah banyak; penuhilah bumi dan taklukkanlah itu, berkuasalah atas ikan-ikan di laut dan burung-burung di udara dan atas segala binatang yang merayap di bumi." Ayat ini sering disalahartikan sebagai izin untuk mengeksploitasi alam habis-habisan. Padahal, kata "berkuasa" (Ibr: radah) dalam konteks Alkitab berarti memimpin dengan penuh tanggung jawab, bukan mendominasi secara destruktif.'),
        body('Mandat pemeliharaan dipertegas dalam Kejadian 2:15: "TUHAN Allah mengambil manusia itu dan menempatkannya dalam taman Eden untuk mengusahakan dan memelihara taman itu." Kata "memelihara" (Ibr: shamar) mengandung makna menjaga, melindungi, merawat, dan bertanggung jawab atas alam. Ini adalah mandat yang diberikan Allah kepada manusia sebelum dosa masuk ke dunia—artinya pemeliharaan alam adalah bagian dari rancangan awal Allah bagi manusia.'),
        body('Mandat budaya (cultural mandate) ini menegaskan bahwa manusia adalah penatalayan (steward) atas ciptaan Allah, bukan tuan yang bisa berbuat sesuka hati. Sebagai penatalayan, manusia akan dimintai pertanggungjawaban atas bagaimana ia mengelola alam yang dipercayakan kepadanya. Yesus sendiri mengajarkan bahwa penatalayan yang baik adalah mereka yang mengelola dengan bijak dan dapat dipercaya (Luk. 16:10-12).'),
        body('Dosa manusia (Kej. 3) menyebabkan hubungan manusia dengan alam menjadi terganggu — alam menjadi "memberontak" dan pengelolaan menjadi lebih sulit. Namun demikian, mandat untuk memelihara alam tidak pernah dicabut. Karya penebusan Kristus justru memulihkan manusia untuk kembali menjalankan mandat tersebut dengan benar.'),
 
        h2('C. Tanggung Jawab Umat Kristen dalam Menjaga Alam'),
        body('Sebagai orang-orang yang telah ditebus dan dipulihkan oleh Kristus, umat Kristen memiliki panggilan istimewa untuk menjadi teladan dalam pemeliharaan alam. Berikut adalah beberapa bentuk tanggung jawab umat Kristen:'),
 
        h3('1. Tanggung Jawab Teologis'),
        body('Tanggung jawab pertama adalah mengakui bahwa alam adalah milik Allah dan manusia adalah penatalayan-Nya. Umat Kristen harus membangun teologi ekologi yang sehat—bahwa mencintai Allah berarti juga mencintai ciptaan-Nya. Ekologi dan iman bukan dua hal yang terpisah, melainkan saling berkaitan erat. Ketidakpedulian terhadap alam adalah bentuk kelalaian rohani.'),
 
        h3('2. Tanggung Jawab Moral dan Etis'),
        body('Umat Kristen harus berani mengambil sikap moral dalam isu-isu lingkungan. Ini mencakup menolak pola hidup konsumtif yang eksploitatif, mendukung kebijakan yang berpihak pada kelestarian lingkungan, serta berani bersuara ketika ada ketidakadilan ekologis—misalnya ketika perusahaan merusak hutan atau mencemari sungai demi keuntungan semata.'),
 
        h3('3. Tanggung Jawab Praktis dan Personal'),
        body('Tanggung jawab ini berkaitan dengan gaya hidup sehari-hari. Umat Kristen dipanggil untuk hidup sederhana, mengurangi pemborosan, menghemat energi, dan membiasakan diri untuk tidak merusak lingkungan. Langkah kecil seperti tidak membuang sampah sembarangan, mendaur ulang, menanam pohon, dan hemat air merupakan wujud konkret dari iman yang hidup.'),
 
        h3('4. Tanggung Jawab Sosial dan Komunal'),
        body('Gereja sebagai komunitas iman memiliki peran strategis dalam mendidik jemaatnya tentang pentingnya pemeliharaan alam. Gereja dapat menginisiasi program-program lingkungan, seperti penghijauan, pengelolaan sampah, penyuluhan lingkungan, serta bersinergi dengan pemerintah dan lembaga masyarakat dalam gerakan pelestarian alam.'),
 
        h2('D. Krisis Lingkungan dan Respons Iman Kristen'),
        body('Dunia saat ini menghadapi berbagai krisis lingkungan yang serius. Pemanasan global mengakibatkan perubahan iklim yang ekstrem, mencairnya es di kutub, dan naiknya permukaan laut. Deforestasi menghancurkan paru-paru bumi dan habitat berbagai spesies yang terancam punah. Pencemaran plastik mengotori lautan dan mengancam kehidupan laut. Krisis air bersih mengancam jutaan orang di seluruh dunia.'),
        body('Iman Kristen tidak boleh berdiam diri menghadapi krisis ini. Sejarah mencatat bahwa gereja selalu hadir di garis terdepan dalam merespons krisis kemanusiaan—baik melalui lembaga pendidikan, kesehatan, maupun keadilan sosial. Sudah saatnya gereja juga hadir secara aktif dalam merespons krisis lingkungan.'),
        body('Beberapa organisasi Kristen internasional, seperti A Rocha (organisasi Kristen pelestarian alam), telah memberikan contoh nyata bagaimana iman dapat mendorong aksi nyata pelestarian lingkungan. Di Indonesia, beberapa gereja dan lembaga Kristen juga mulai mengintegrasikan kepedulian lingkungan dalam program pelayanan mereka.'),
        body('Elia dalam 1 Raja-raja 19 mengalami kehabisan semangat di padang gurun, namun Allah hadir dan memulihkannya melalui alam. Ini mengingatkan kita bahwa alam adalah ruang di mana Allah hadir dan berbicara. Menghancurkan alam berarti menutup salah satu kanal komunikasi Allah dengan manusia.'),
 
        h2('E. Wujud Nyata Tanggung Jawab Umat Kristen'),
        body('Tanggung jawab umat Kristen dalam menjaga alam tidak boleh berhenti pada tataran wacana dan teori semata. Iman yang sejati harus diwujudnyatakan dalam tindakan konkret (Yak. 2:17). Berikut adalah wujud nyata yang dapat dilakukan:'),
 
        new Paragraph({
          numbering: { reference: 'bullets', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 80 },
          children: [new TextRun({ text: 'Menerapkan gaya hidup ramah lingkungan: menggunakan produk yang dapat didaur ulang, mengurangi penggunaan plastik sekali pakai, dan menghemat listrik serta air.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          numbering: { reference: 'bullets', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 80 },
          children: [new TextRun({ text: 'Aktif dalam kegiatan penanaman pohon dan penghijauan lingkungan sekitar, baik secara pribadi maupun bersama komunitas gereja.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          numbering: { reference: 'bullets', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 80 },
          children: [new TextRun({ text: 'Mendukung pertanian organik dan ramah lingkungan sebagai bentuk pengelolaan alam yang bertanggung jawab.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          numbering: { reference: 'bullets', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 80 },
          children: [new TextRun({ text: 'Mendidik generasi muda dalam gereja tentang pentingnya pelestarian alam melalui kelas sekolah minggu, persekutuan pemuda, dan retret alam.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          numbering: { reference: 'bullets', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 80 },
          children: [new TextRun({ text: 'Berdoa bagi pemulihan alam — menjadikan doa untuk ciptaan Allah sebagai bagian dari ibadah jemaat secara rutin.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          numbering: { reference: 'bullets', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 80 },
          children: [new TextRun({ text: 'Berpartisipasi aktif dalam advokasi kebijakan lingkungan yang berpihak pada kelestarian alam dan keadilan bagi masyarakat yang terdampak kerusakan lingkungan.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          numbering: { reference: 'bullets', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 200 },
          children: [new TextRun({ text: 'Mengelola bangunan gereja secara ramah lingkungan: menggunakan panel surya, sistem penampung air hujan, dan taman hijau di sekitar gereja.', size: 24, font: 'Times New Roman' })]
        }),
 
        body('Semua tindakan nyata ini mencerminkan bahwa iman Kristen bukan hanya urusan akhirat, melainkan juga memiliki relevansi dan dampak nyata bagi kehidupan di dunia sekarang ini—termasuk bagi kelestarian alam semesta yang dipercayakan Allah kepada umat-Nya.'),
      ]
    },
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1701, right: 1134, bottom: 1134, left: 1701 }
        }
      },
      children: [
        h1('BAB III'),
        h1('PENUTUP'),
 
        h2('A. Kesimpulan'),
        body('Berdasarkan pembahasan yang telah diuraikan dalam makalah ini, dapat ditarik beberapa kesimpulan sebagai berikut:'),
        new Paragraph({
          numbering: { reference: 'numbers3', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 120 },
          children: [new TextRun({ text: 'Alam semesta adalah ciptaan Allah yang baik dan mulia. Seluruh ciptaan mencerminkan kemuliaan Allah dan memiliki nilai di hadapan-Nya. Alam bukan milik manusia, melainkan milik Allah yang dipercayakan kepada manusia untuk dikelola dengan bijaksana.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          numbering: { reference: 'numbers3', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 120 },
          children: [new TextRun({ text: 'Allah memberikan mandat kepada manusia—dan secara khusus kepada umat Kristen—untuk mengusahakan dan memelihara alam (Kej. 2:15). Mandat ini tidak pernah dicabut dan tetap berlaku hingga hari ini. Sebagai penatalayan Allah, manusia bertanggung jawab untuk mengelola alam dengan penuh integritas dan kebijaksanaan.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          numbering: { reference: 'numbers3', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 120 },
          children: [new TextRun({ text: 'Umat Kristen memiliki tanggung jawab teologis, moral, praktis, dan sosial dalam menjaga alam. Tanggung jawab ini harus diwujudkan dalam tindakan nyata—baik pada level pribadi, keluarga, gereja, maupun masyarakat luas.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          numbering: { reference: 'numbers3', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 120 },
          children: [new TextRun({ text: 'Krisis lingkungan yang terjadi saat ini merupakan panggilan mendesak bagi gereja untuk bangkit dan hadir sebagai agen pemulihan—tidak hanya bagi jiwa manusia, tetapi juga bagi ciptaan Allah yang merintih menantikan pemulihan (Rm. 8:22).', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          numbering: { reference: 'numbers3', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 200 },
          children: [new TextRun({ text: 'Menjaga alam adalah bentuk ibadah kepada Allah. Setiap tindakan pelestarian lingkungan yang dilakukan dengan motivasi iman kepada Allah adalah persembahan yang harum di hadapan-Nya.', size: 24, font: 'Times New Roman' })]
        }),
 
        h2('B. Saran'),
        body('Berdasarkan kesimpulan di atas, beberapa saran dapat dikemukakan:'),
        new Paragraph({
          numbering: { reference: 'numbers4', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 120 },
          children: [new TextRun({ text: 'Gereja dan lembaga Kristen perlu mengintegrasikan tema pemeliharaan alam ke dalam kurikulum pendidikan agama, khotbah, dan program pelayanan secara sistematis dan berkelanjutan.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          numbering: { reference: 'numbers4', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 120 },
          children: [new TextRun({ text: 'Setiap orang Kristen hendaknya merefleksikan gaya hidupnya dan melakukan perubahan konkret menuju gaya hidup yang lebih ramah lingkungan, sesuai dengan panggilan iman sebagai penatalayan ciptaan Allah.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          numbering: { reference: 'numbers4', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 120 },
          children: [new TextRun({ text: 'Para pemimpin gereja perlu mengambil peran aktif dalam advokasi kebijakan lingkungan dan membangun koalisi lintas denominasi bahkan lintas agama dalam upaya bersama menjaga kelestarian alam.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          numbering: { reference: 'numbers4', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 120 },
          children: [new TextRun({ text: 'Generasi muda Kristen perlu didorong untuk terlibat aktif dalam gerakan pelestarian lingkungan, termasuk melalui pendidikan formal, komunitas peduli alam, dan pemanfaatan media sosial untuk menyebarkan kesadaran ekologis yang berbasis iman.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          numbering: { reference: 'numbers4', level: 0 },
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 200 },
          children: [new TextRun({ text: 'Penelitian dan kajian teologi ekologi perlu terus dikembangkan agar iman Kristen semakin relevan dalam menjawab tantangan lingkungan yang dihadapi umat manusia saat ini dan masa depan.', size: 24, font: 'Times New Roman' })]
        }),
 
        ...empty(1),
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: 360, lineRule: 'auto', after: 120 },
          children: [
            new TextRun({ text: 'Marilah kita, sebagai umat Kristen, menjawab panggilan mulia ini dengan sungguh-sungguh. Karena menjaga alam bukan sekadar kepedulian ekologis, melainkan ungkapan kasih kita kepada Allah Sang Pencipta yang telah memberikan alam ini sebagai anugerah dan amanah yang tak ternilai harganya.', size: 24, font: 'Times New Roman', italics: true })
          ]
        }),
 
        ...empty(2),
        // DAFTAR PUSTAKA
        h1('DAFTAR PUSTAKA'),
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          indent: { hanging: 720, left: 720 },
          spacing: { after: 120, line: 360, lineRule: 'auto' },
          children: [new TextRun({ text: 'Alkitab Terjemahan Baru. (2008). Jakarta: Lembaga Alkitab Indonesia.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          indent: { hanging: 720, left: 720 },
          spacing: { after: 120, line: 360, lineRule: 'auto' },
          children: [new TextRun({ text: 'Berry, R. J. (2006). ', size: 24, font: 'Times New Roman' }), new TextRun({ text: 'Care for Creation: A Christian Response to the Environmental Crisis.', size: 24, font: 'Times New Roman', italics: true }), new TextRun({ text: ' Nottingham: Inter-Varsity Press.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          indent: { hanging: 720, left: 720 },
          spacing: { after: 120, line: 360, lineRule: 'auto' },
          children: [new TextRun({ text: 'Bouma-Prediger, S. (2001). ', size: 24, font: 'Times New Roman' }), new TextRun({ text: 'For the Beauty of the Earth: A Christian Vision for Creation Care.', size: 24, font: 'Times New Roman', italics: true }), new TextRun({ text: ' Grand Rapids: Baker Academic.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          indent: { hanging: 720, left: 720 },
          spacing: { after: 120, line: 360, lineRule: 'auto' },
          children: [new TextRun({ text: 'DeWitt, C. B. (1998). ', size: 24, font: 'Times New Roman' }), new TextRun({ text: 'Caring for Creation: Responsible Stewardship of God\'s Handiwork.', size: 24, font: 'Times New Roman', italics: true }), new TextRun({ text: ' Grand Rapids: Baker Books.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          indent: { hanging: 720, left: 720 },
          spacing: { after: 120, line: 360, lineRule: 'auto' },
          children: [new TextRun({ text: 'Hasel, G. F. (1991). ', size: 24, font: 'Times New Roman' }), new TextRun({ text: 'Old Testament Theology: Basic Issues in the Current Debate.', size: 24, font: 'Times New Roman', italics: true }), new TextRun({ text: ' Grand Rapids: Eerdmans.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          indent: { hanging: 720, left: 720 },
          spacing: { after: 120, line: 360, lineRule: 'auto' },
          children: [new TextRun({ text: 'Moltmann, J. (1985). ', size: 24, font: 'Times New Roman' }), new TextRun({ text: 'God in Creation: A New Theology of Creation and the Spirit of God.', size: 24, font: 'Times New Roman', italics: true }), new TextRun({ text: ' San Francisco: Harper & Row.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          indent: { hanging: 720, left: 720 },
          spacing: { after: 120, line: 360, lineRule: 'auto' },
          children: [new TextRun({ text: 'Sirait, M. T. (2015). Gereja dan Tanggung Jawab Lingkungan. ', size: 24, font: 'Times New Roman' }), new TextRun({ text: 'Jurnal Teologi Reformed Indonesia,', size: 24, font: 'Times New Roman', italics: true }), new TextRun({ text: ' 5(2), 45-62.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          indent: { hanging: 720, left: 720 },
          spacing: { after: 120, line: 360, lineRule: 'auto' },
          children: [new TextRun({ text: 'Stott, J. R. W. (1984). ', size: 24, font: 'Times New Roman' }), new TextRun({ text: 'Issues Facing Christians Today.', size: 24, font: 'Times New Roman', italics: true }), new TextRun({ text: ' Basingstoke: Marshall Pickering.', size: 24, font: 'Times New Roman' })]
        }),
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          indent: { hanging: 720, left: 720 },
          spacing: { after: 120, line: 360, lineRule: 'auto' },
          children: [new TextRun({ text: 'Wright, C. J. H. (2006). ', size: 24, font: 'Times New Roman' }), new TextRun({ text: 'The Mission of God: Unlocking the Bible\'s Grand Narrative.', size: 24, font: 'Times New Roman', italics: true }), new TextRun({ text: ' Downers Grove: IVP Academic.', size: 24, font: 'Times New Roman' })]
        }),
      ]
    }
  ],
  numbering: {
    config: [
      {
        reference: 'bullets',
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: 'numbers',
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: 'numbers2',
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: 'numbers3',
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: 'numbers4',
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  }
});
 
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('C:/Users/Martin/Downloads/tugas/tugas.docx', buffer);
  console.log('Done! Document created with automatic Table of Contents.');
  console.log('Note: When you open the document in Word, it will ask to update fields. Click "Yes" to update the TOC.');
}).catch(err => console.error(err));

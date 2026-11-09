// Сборка отчёта по практической работе №1 в формате DOCX.
// Запуск: [VARIANT=src/variantN.js] node src/build.js <выходной.docx> [layout.json]
// layout.json (необязательно) содержит номера страниц для оглавления
// и разбиение Таблицы 1 по страницам; формируется скриптом src/make.py.
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, ShadingType, Footer, PageNumber,
  LevelFormat, TabStopType, VerticalAlign, HeadingLevel, Tab,
} = require('docx');

const ROOT = path.resolve(__dirname, '..');
const OUT = process.argv[2] || path.join(ROOT, 'report.docx');
const layout = process.argv[3] && fs.existsSync(process.argv[3])
  ? JSON.parse(fs.readFileSync(process.argv[3], 'utf8')) : {};

// ---------- Единицы и константы макета ----------
const CM = 567;                       // 1 см в DXA (twips)
const TEXT_W = 11906 - 3 * CM - 850;  // A4 минус поля 3 см и 1,5 см = 9355
const FONT = 'Times New Roman';
const LINE15 = { line: 360, lineRule: 'auto', before: 0, after: 0 };
const LINE10 = { line: 240, lineRule: 'auto', before: 0, after: 0 };
const INDENT = { firstLine: 709 };    // красная строка 1,25 см

// ---------- Базовые элементы ----------
const run = (text, o = {}) => new TextRun({ text, font: FONT, size: 28, ...o });

// Абзац основного текста: 14 пт, по ширине, красная строка 1,25 см, интервал 1,5
function P(content, o = {}) {
  const children = typeof content === 'string' ? [run(content)] : content;
  return new Paragraph({
    children, alignment: AlignmentType.JUSTIFIED, spacing: LINE15, indent: INDENT, ...o,
  });
}
const EMPTY = () => new Paragraph({ children: [run('')], spacing: LINE15 });

// Заголовок: 16 пт, полужирный. Ненумерованный – по центру без отступа,
// нумерованный – по ширине с отступом 1,25 см. Новый заголовок – с новой страницы.
function H(text, { numbered = false, level = 1, newPage = true } = {}) {
  return new Paragraph({
    heading: level === 1 ? HeadingLevel.HEADING_1 : HeadingLevel.HEADING_2,
    children: [new TextRun({ text, font: FONT, size: 32, bold: true })],
    alignment: numbered ? AlignmentType.JUSTIFIED : AlignmentType.CENTER,
    indent: numbered ? INDENT : { firstLine: 0 },
    spacing: LINE15, keepNext: true, pageBreakBefore: newPage,
  });
}

// Списки (нумерованный – арабские цифры, маркированный – символ по умолчанию)
let listInstance = 0;
const newList = () => ++listInstance;
const LI = (text, ref, inst) => new Paragraph({
  children: typeof text === 'string' ? [run(text)] : text,
  numbering: { reference: ref, level: 0, instance: inst },
  alignment: AlignmentType.JUSTIFIED, spacing: LINE15,
});
const numList = (items) => { const i = newList(); return items.map((t) => LI(t, 'num', i)); };
const bulList = (items) => { const i = newList(); return items.map((t) => LI(t, 'bul', i)); };

// Подпись таблицы: 12 пт, слева сверху, без отступа
const tableLabel = (text, o = {}) => new Paragraph({
  children: [new TextRun({ text, font: FONT, size: 24 })],
  alignment: AlignmentType.LEFT, spacing: LINE10, indent: { firstLine: 0 }, keepNext: true, ...o,
});

// ---------- Таблицы ----------
const BORDER = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };
const cellPara = (text, { bold = false, center = false } = {}) => new Paragraph({
  children: [new TextRun({ text, font: FONT, size: 24, bold })],
  alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
  indent: { firstLine: 0 }, spacing: LINE10,
});
function cell(content, width, { bold = false, center = false, header = false } = {}) {
  const parts = Array.isArray(content) ? content : [content];
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: BORDERS,
    margins: { top: 60, bottom: 60, left: 90, right: 90 },
    verticalAlign: header || bold ? VerticalAlign.CENTER : VerticalAlign.TOP,
    shading: header ? { fill: 'F2F2F2', type: ShadingType.CLEAR, color: 'auto' } : undefined,
    children: parts.map((t) => cellPara(t, { bold, center })),
  });
}
// Перекрёстная таблица: заголовки сверху и в левом столбце – полужирные, по центру
function crossTable(headers, rows, widths) {
  const head = new TableRow({
    tableHeader: true, cantSplit: true,
    children: headers.map((h, i) => cell(h, widths[i], { bold: true, center: true, header: true })),
  });
  const body = rows.map((r) => new TableRow({
    cantSplit: true,
    children: r.map((c, i) => cell(c, widths[i], i === 0 ? { bold: true, center: true } : {})),
  }));
  return new Table({
    width: { size: TEXT_W, type: WidthType.DXA }, columnWidths: widths, rows: [head, ...body],
  });
}
// Таблица на нескольких страницах: шапка дублируется на каждой странице,
// со второй страницы (если таблица занимает 3 и более страниц) – «Продолжение Таблицы N»,
// в конце (если таблица занимает 2 и более страниц) – «Конец Таблицы N».
// chunks – число строк таблицы на каждой странице (определяется src/make.py).
const tableMarkers = {};
function multiPageTable(num, caption, headers, rows, widths, chunks) {
  tableMarkers[num] = rows.map((r) => r[1]);
  const out = [tableLabel(`Таблица ${num} – ${caption}`)];
  if (!chunks || chunks.length < 2) {
    out.push(crossTable(headers, rows, widths), EMPTY());
    return out;
  }
  if (chunks.length === 2) {
    out.push(crossTable(headers, rows, widths));
  } else {
    let start = 0;
    chunks.forEach((n, k) => {
      if (k > 0) out.push(tableLabel(`Продолжение Таблицы ${num}`, { pageBreakBefore: true }));
      out.push(crossTable(headers, rows.slice(start, start + n), widths));
      start += n;
    });
  }
  out.push(tableLabel(`Конец Таблицы ${num}`, { keepNext: false }), EMPTY());
  return out;
}

// ---------- Содержимое ----------
const VARIANT = process.env.VARIANT || path.join(__dirname, 'variant1.js');
const content = require(path.resolve(VARIANT))({ H, P, EMPTY, run, numList, bulList, multiPageTable, layout });
const { TOPIC, TOC_ITEMS } = content;

// Титульный лист (по шаблону)
const titleLine = (text, o = {}) => new Paragraph({
  children: [run(text, o.run || {})], alignment: o.align || AlignmentType.CENTER,
  spacing: LINE15, indent: { firstLine: 0 },
});
const right = (t) => titleLine(t, { align: AlignmentType.RIGHT });
const titlePage = [
  ...Array(5).fill(0).map(EMPTY),
  titleLine('ПРАКТИЧЕСКАЯ РАБОТА №1', { run: { bold: true, size: 32 } }),
  EMPTY(),
  titleLine(`Тема: «${TOPIC}»`, { run: { bold: true } }),
  ...Array(5).fill(0).map(EMPTY),
  right('Дисциплина: …'), EMPTY(),
  right('Преподаватель: Семак Вячеслав Павлович'), EMPTY(),
  right('Выполнил: …'), EMPTY(),
  right('Группа: …'), EMPTY(),
  right('Дата выполнения: «__»_____ 20__ г.'),
];
// Год размещается внизу титульного листа (в колонтитуле первой страницы, без номера страницы)
const titleFooter = new Footer({ children: [titleLine('2026')] });

// Оглавление (иерархическая лестница; полужирным выделен только заголовок «Оглавление»)
const pages = layout.pages || {};
const toc = [
  new Paragraph({
    children: [new TextRun({ text: 'Оглавление', font: FONT, size: 32, bold: true })],
    alignment: AlignmentType.CENTER, spacing: LINE15, pageBreakBefore: true,
  }),
  EMPTY(),
  ...TOC_ITEMS.map(([t, lvl]) => new Paragraph({
    children: [run(t), new TextRun({ children: [new Tab()], font: FONT, size: 28 }), run(String(pages[t] || 0))],
    tabStops: [{ type: TabStopType.RIGHT, position: TEXT_W, leader: 'dot' }],
    indent: { left: lvl * 709, right: 0, firstLine: 0 },
    alignment: AlignmentType.LEFT, spacing: LINE15,
  })),
];

// ---------- Документ ----------
const listLevel = (format, text) => ({
  level: 0, format, text, alignment: AlignmentType.LEFT,
  style: {
    paragraph: { indent: { left: 0, firstLine: 709 }, tabStops: [{ type: TabStopType.LEFT, position: 1134 }] },
    run: { font: FONT, size: 28 },
  },
});

const pageFooter = new Footer({
  children: [new Paragraph({
    alignment: AlignmentType.CENTER, indent: { firstLine: 0 }, spacing: LINE10,
    children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 24 })],
  })],
});

const doc = new Document({
  creator: 'Студент',
  title: `Практическая работа №1. ${TOPIC}`,
  styles: {
    default: { document: { run: { font: FONT, size: 28, color: '000000' }, paragraph: { spacing: LINE15 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 32, bold: true, color: '000000' },
        paragraph: { spacing: LINE15, keepNext: true, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 32, bold: true, color: '000000' },
        paragraph: { spacing: LINE15, keepNext: true, outlineLevel: 1 } },
    ],
  },
  numbering: {
    config: [
      { reference: 'num', levels: [listLevel(LevelFormat.DECIMAL, '%1.')] },
      { reference: 'bul', levels: [listLevel(LevelFormat.BULLET, '•')] },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 2 * CM, bottom: 2 * CM, left: 3 * CM, right: 850, footer: 567 },
      },
      titlePage: true,
    },
    headers: {},
    footers: { default: pageFooter, first: titleFooter },
    children: [
      ...titlePage, ...toc, ...content.body,
    ],
  }],
});

if (process.env.MARKERS) {
  fs.writeFileSync(process.env.MARKERS, JSON.stringify({ tables: tableMarkers, toc: TOC_ITEMS.map((t) => t[0]) }));
}
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log('written', OUT); });

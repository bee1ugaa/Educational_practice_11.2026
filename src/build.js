// Сборка отчёта по практической работе №1 в формате DOCX.
// Запуск: node src/build.js <выходной.docx> [layout.json]
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
const TOPIC = 'Сравнение архитектуры операционных систем Windows NT, Linux и MINIX 3';
const TAB1 = 'Сравнительная таблица операционных систем';
const TAB2 = 'Последствия компрометации драйвера для конфиденциальности, целостности и доступности';

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
const TOC_ITEMS = [
  ['Введение', 0], ['Цель и задачи работы', 0], ['Средства выполнения работы', 0],
  ['Ход работы', 0],
  ['1. Рассматриваемые версии операционных систем', 1],
  ['2. Архитектура Windows NT', 1],
  ['3. Архитектура ОС на базе Linux', 1],
  ['4. Архитектура MINIX 3', 1],
  ['5. Сравнительный анализ архитектур', 1],
  ['6. Влияние размещения компонентов на изоляцию, устойчивость и безопасность', 1],
  ['7. Разбор ситуации компрометации драйвера', 1],
  ['Вывод', 0], ['Список использованных источников', 0],
];
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

const intro = [
  H('Введение'), EMPTY(),
  P('Ядро операционной системы (ОС) выполняется в привилегированном режиме процессора и имеет полный доступ к оборудованию. Архитектура ОС определяет, какие компоненты работают в пространстве ядра, а какие вынесены в пространство пользователя. Различают монолитные ядра, микроядра и гибридные ядра, сочетающие оба подхода.'),
  P('От архитектуры зависят производительность, устойчивость к отказам и информационная безопасность: ошибка в компоненте ядра способна нарушить работу всей системы, а сбой изолированного компонента затрагивает только его самого. В работе рассматриваются Windows NT (гибридное ядро), Linux (монолитное модульное ядро) и MINIX 3 (микроядро).'),
];

const goals = [
  H('Цель и задачи работы'), EMPTY(),
  P([run('Цель работы:', { bold: true })]),
  P('Получение практических и теоретических навыков в сравнительном анализе архитектуры операционных систем. В данной практической работе необходимо произвести сравнение архитектуры Windows NT, Linux и MINIX 3, определить расположение основных системных компонентов и проанализировать значение архитектурных решений для изоляции, устойчивости и информационной безопасности.'),
  P([run('Задачи для выполнения работы:', { bold: true })]),
  ...numList([
    'Изучить техническую документацию по архитектуре Windows NT, Linux и MINIX 3, указав рассматриваемые версии систем и использованные источники.',
    'Определить тип ядра каждой операционной системы и выделить особенности её архитектурной организации.',
    'Установить назначение и расположение основных компонентов ОС: ядра, системных служб, драйверов, компонентов файловой системы, средств управления процессами и памятью.',
    'Рассмотреть способы взаимодействия приложений, системных служб, драйверов и ядра, обозначив границы пространства пользователя и пространства ядра.',
    'Составить упрощённые архитектурные схемы Windows NT, Linux и MINIX 3, отразив размещение компонентов и связи между ними.',
    'Выполнить сравнительный анализ трёх ОС по единым критериям и представить результаты в сравнительной таблице.',
    'Проанализировать влияние размещения и полномочий системных компонентов на их изоляцию, устойчивость ОС и информационную безопасность.',
    'Рассмотреть ситуацию компрометации драйвера, определить возможные последствия для конфиденциальности, целостности и доступности и предложить меры снижения риска.',
    'Сформулировать обоснованные выводы о сходствах, различиях и ограничениях исследованных архитектур.',
    'Оформить отчёт, содержащий архитектурные схемы, сравнительную таблицу, разбор ситуации, выводы и список источников.',
  ]),
];

const tools = [
  H('Средства выполнения работы'), EMPTY(),
  P('Для выполнения работы были использованы следующие средства:'),
  ...bulList([
    'технические: персональный компьютер с доступом к сети Интернет;',
    'программные: браузер, текстовый редактор, средство просмотра PDF;',
    'информационные: документация Microsoft [1–4], документация ядра Linux и Ubuntu [5–7], материалы проекта MINIX 3 [8–10], учебная литература [11].',
  ]),
];

const ch1 = [
  H('Ход работы'), EMPTY(),
  H('1. Рассматриваемые версии операционных систем', { numbered: true, level: 2, newPage: false }), EMPTY(),
  P('Для анализа были выбраны следующие версии ОС:'),
  ...bulList([
    'Windows 11 версии 24H2 (ядро NT 10.0), x86-64 [1–4];',
    'Ubuntu 24.04 LTS с ядром Linux 6.8, x86-64 [5–7];',
    'MINIX 3.4.0, x86 [8–10].',
  ]),
  P('Все три системы были рассмотрены по единому набору критериев, приведённому в сравнительной таблице (раздел 5).'),
];

const ch2 = [
  H('2. Архитектура Windows NT', { numbered: true, level: 2 }), EMPTY(),
  P('Ядро Windows NT относится к гибридному типу. Структура NT повторяет микроядерную модель (ядро, исполнительная система, подсистемы окружения), но ради производительности исполнительная система, драйверы и графическая подсистема Win32k.sys выполняются в режиме ядра в общем адресном пространстве [1, 2].'),
  P('В режиме ядра находятся уровень аппаратных абстракций HAL, ядро (планирование потоков, прерывания), исполнительная система ntoskrnl.exe (диспетчеры объектов, ввода-вывода, памяти, процессов, монитор безопасности, реестр, ALPC) и драйверы. В пространстве пользователя работают приложения, системные процессы (smss, csrss, winlogon, lsass), службы, библиотеки подсистем и ntdll.dll, а также драйверы UMDF [3].'),
  P('Приложение вызывает Win32 API, библиотека обращается к ntdll.dll, которая инструкцией syscall передаёт управление ядру; ввод-вывод выполняется пакетами IRP через стек драйверов.'),
  P('Компоненты режима ядра не изолированы друг от друга. Дополнительную изоляцию даёт технология VBS: гипервизор выделяет защищённую среду для Secure Kernel и процесса LsaIso.exe [1].'),
];

const ch3 = [
  H('3. Архитектура ОС на базе Linux', { numbered: true, level: 2 }), EMPTY(),
  P('Ядро Linux монолитное: планировщик, управление памятью, VFS, файловые системы, сетевой стек и драйверы работают в едином адресном пространстве и вызывают функции друг друга напрямую. Драйверы могут подключаться как загружаемые модули (*.ko), но после загрузки получают полные полномочия ядра [5].'),
  P('В ядре 6.8 используется планировщик EEVDF, страничная виртуальная память с копированием при записи, cgroups и namespaces для изоляции групп процессов, модули безопасности LSM (AppArmor в Ubuntu). В пространстве пользователя работают приложения, glibc, systemd и системные службы; отдельные драйверы и ФС могут выполняться там же через FUSE, UIO и VFIO [5, 7].'),
  P('Граница привилегий проходит по интерфейсу системных вызовов (а также /proc, /sys, ioctl и netlink): всё, что расположено ниже неё, включая сторонние модули, работает в нулевом кольце защиты.'),
];

const ch4 = [
  H('4. Архитектура MINIX 3', { numbered: true, level: 2 }), EMPTY(),
  P('MINIX 3 построена на микроядре объёмом несколько тысяч строк кода, которое обрабатывает прерывания, переключает процессы и передаёт сообщения. В ядре также работают задачи SYSTEM и CLOCK [8, 9].'),
  P('Остальные компоненты – отдельные процессы пространства пользователя: драйверы устройств, серверы PM (процессы), VM (память), SCHED (планирование), VFS и MFS (файловые системы), INET (сеть), DS (хранилище данных), RS (сервер реинкарнации) и пользовательские процессы [8, 10].'),
  P('Компоненты взаимодействуют только сообщениями через ядро (send, receive, sendrec, notify). Права процесса ограничены маской привилегий, а доступ к чужой памяти возможен только через гранты.'),
  P('Сервер реинкарнации RS контролирует драйверы и серверы и автоматически перезапускает отказавший компонент; для блочных драйверов восстановление прозрачно для приложений [9].'),
];

// Таблица 1
const T1_HEAD = ['Критерий', 'Windows NT', 'ОС на базе Linux', 'MINIX 3'];
const T1_W = [2401, 2318, 2318, 2318];
const T1_ROWS = [
  ['Рассматриваемая версия ОС', 'Windows 11 24H2 (NT 10.0)', 'Ubuntu 24.04 LTS, ядро 6.8', 'MINIX 3.4.0'],
  ['Название и тип ядра', 'ntoskrnl.exe; гибридное', 'Linux; монолитное модульное', 'Микроядро MINIX 3'],
  ['Основные компоненты пространства ядра',
    'HAL, ядро, исполнительная система, Win32k.sys, драйверы и ФС',
    'Планировщик, память, VFS и ФС, сетевой стек, LSM, драйверы и модули',
    'Микроядро, задачи SYSTEM и CLOCK'],
  ['Основные компоненты пространства пользователя',
    'Приложения, системные процессы, службы, DLL подсистем, ntdll.dll, UMDF',
    'Приложения, glibc, systemd и службы, FUSE',
    'Приложения, серверы PM, VM, VFS, RS и др., все драйверы'],
  ['Взаимодействие приложений, служб и ядра',
    'Win32 API → ntdll.dll → syscall; ALPC, RPC; IRP',
    'Системные вызовы, /proc, /sys, ioctl, netlink',
    'Сообщения через ядро; вызовы ядра по маске привилегий'],
  ['Размещение драйверов устройств', 'В ядре (WDM, KMDF); часть – UMDF', 'В ядре (модули *.ko); часть – FUSE, VFIO', 'В пространстве пользователя, каждый – отдельный процесс'],
  ['Размещение файловых служб', 'Драйверы ФС (ntfs.sys) в ядре', 'VFS и ФС (ext4) в ядре', 'Серверы VFS и MFS в пространстве пользователя'],
  ['Управление процессами и памятью – основные особенности',
    'Приоритетное вытесняющее планирование потоков; страничная память',
    'Планировщик EEVDF, cgroups; страничная память, копирование при записи',
    'Механизм в ядре, политика в серверах PM, SCHED, VM'],
  ['Способы изоляции системных компонентов',
    'Кольца 0 и 3, адресные пространства процессов, VBS; внутри ядра изоляции нет',
    'Кольца 0 и 3, namespaces, seccomp, LSM; внутри ядра изоляции нет',
    'Отдельное адресное пространство и маска привилегий у каждого компонента'],
  ['Возможные последствия сбоя драйвера', 'Критическая ошибка («синий экран»), перезагрузка', 'Oops или kernel panic', 'Завершается только процесс драйвера'],
  ['Средства восстановления после отказа компонента',
    'Автоперезагрузка, дампы памяти, WinRE, перезапуск служб',
    'kdump, panic=N, перезапуск служб systemd',
    'Автоматический перезапуск компонента сервером RS'],
  ['Архитектурные особенности, значимые для безопасности',
    'Подпись драйверов, PatchGuard, HVCI; большая TCB',
    'Подпись модулей, lockdown, KASLR; большая TCB',
    'Минимальная TCB, наименьшие привилегии; ниже производительность'],
];

const ch5 = [
  H('5. Сравнительный анализ архитектур', { numbered: true, level: 2 }), EMPTY(),
  P('Результаты сравнения трёх ОС по единым критериям сведены в таблицу (Таблица 1 – Сравнительная таблица операционных систем).'),
  ...multiPageTable(1, TAB1, T1_HEAD, T1_ROWS, T1_W, layout.table1),
  P('Во всех трёх системах используются разделение режимов процессора и изоляция адресных пространств процессов. Главное различие – объём кода с максимальными привилегиями: в Windows NT и Linux драйверы, файловые системы и сетевой стек находятся в ядре, в MINIX 3 – в отдельных непривилегированных процессах.'),
];

const ch6 = [
  H('6. Влияние размещения компонентов на изоляцию, устойчивость и безопасность', { numbered: true, level: 2 }), EMPTY(),
  P('Компонент в пространстве ядра имеет доступ ко всей памяти и оборудованию, и последствия его ошибки ничем не ограничены. Поэтому архитектура влияет на систему следующим образом:'),
  ...bulList([
    'изоляция: в Windows NT и Linux граница проходит только между пользователем и ядром, в MINIX 3 – вокруг каждого компонента;',
    'устойчивость: отказ драйвера в Windows NT и Linux обычно приводит к отказу всей системы, в MINIX 3 – устраняется перезапуском драйвера;',
    'безопасность: уязвимость в любом драйвере Windows или модуле Linux даёт привилегии ядра;',
    'производительность: общее адресное пространство избавляет от накладных расходов на сообщения, поэтому Windows NT и Linux быстрее MINIX 3.',
  ]),
  P('Windows и Linux компенсируют эти недостатки подписью драйверов и модулей, HVCI, lockdown и KASLR [4, 6], а также выносом части драйверов в пространство пользователя (UMDF, FUSE).'),
];

// Таблица 2
const T2_HEAD = ['Свойство информации', 'Windows NT', 'ОС на базе Linux', 'MINIX 3'];
const T2_W = [2401, 2318, 2318, 2318];
const T2_ROWS = [
  ['Конфиденциальность', 'Чтение памяти ядра и процессов, включая учётные данные lsass.exe', 'Чтение памяти любых процессов и ключей шифрования', 'Доступ только к данным самого драйвера'],
  ['Целостность', 'Изменение структур ядра, отключение защиты, руткит', 'Изменение структур ядра, отключение LSM и аудита', 'Искажение данных драйвера; ядро и другие компоненты защищены'],
  ['Доступность', 'Отказ всей системы', 'Kernel panic, отказ всей системы', 'Отказ одного устройства до перезапуска драйвера'],
];

const ch7 = [
  H('7. Разбор ситуации компрометации драйвера', { numbered: true, level: 2 }), EMPTY(),
  P('Рассмотрена ситуация, в которой нарушитель с правами администратора загружает подписанный, но уязвимый драйвер (атака BYOVD) и выполняет произвольный код с его полномочиями. Последствия приведены в таблице (Таблица 2 – Последствия компрометации драйвера для конфиденциальности, целостности и доступности).'),
  ...multiPageTable(2, TAB2, T2_HEAD, T2_ROWS, T2_W, layout.table2),
  P('В Windows NT и Linux компрометация драйвера равнозначна компрометации всего узла, в MINIX 3 ущерб ограничен одним драйвером. Для снижения риска предлагаются следующие меры:'),
  ...numList([
    'допускать только подписанные драйверы и блокировать известные уязвимые драйверы [4, 6];',
    'включить Secure Boot, а в Windows – VBS и HVCI;',
    'удалить неиспользуемые драйверы и модули, применять IOMMU;',
    'своевременно устанавливать обновления и ограничить круг администраторов;',
    'контролировать загрузку драйверов и модулей средствами аудита.',
  ]),
];

const conclusion = [
  H('Вывод'), EMPTY(),
  P('В ходе выполнения практической работы были получены практические и теоретические навыки в сравнительном анализе архитектуры операционных систем Windows NT, Linux и MINIX 3.'),
  P('Было установлено, что Windows NT использует гибридное ядро, Linux – монолитное модульное, MINIX 3 – микроядро. Общими являются разделение режимов процессора и изоляция процессов, различие – в объёме кода, выполняемого в ядре. В Windows NT и Linux компрометация драйвера ведёт к компрометации всего узла, в MINIX 3 – ограничена одним драйвером. Монолитные и гибридные ядра выигрывают в производительности, но требуют дополнительных мер защиты, а микроядро обеспечивает лучшую изоляцию ценой скорости работы.'),
];

const sources = [
  H('Список использованных источников'), EMPTY(),
  ...numList([
    'Yosifovich P., Ionescu A., Russinovich M. E., Solomon D. A. Windows Internals. Part 1. – 7th ed. – Redmond : Microsoft Press, 2017. – 800 p.',
    'Overview of Windows Components [Электронный ресурс] // Microsoft Learn. – URL: https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/overview-of-windows-components (дата обращения: 09.11.2026).',
    'User mode and kernel mode [Электронный ресурс] // Microsoft Learn. – URL: https://learn.microsoft.com/en-us/windows-hardware/drivers/gettingstarted/user-mode-and-kernel-mode (дата обращения: 09.11.2026).',
    'Microsoft recommended driver block rules [Электронный ресурс] // Microsoft Learn. – URL: https://learn.microsoft.com/en-us/windows/security/application-security/application-control/app-control-for-business/design/microsoft-recommended-driver-block-rules (дата обращения: 09.11.2026).',
    'The Linux Kernel documentation [Электронный ресурс]. – URL: https://docs.kernel.org/ (дата обращения: 09.11.2026).',
    'Kernel module signing facility [Электронный ресурс] // The Linux Kernel documentation. – URL: https://docs.kernel.org/admin-guide/module-signing.html (дата обращения: 09.11.2026).',
    'Ubuntu 24.04 LTS (Noble Numbat) release notes [Электронный ресурс] // Ubuntu Community Hub. – URL: https://discourse.ubuntu.com/t/ubuntu-24-04-lts-noble-numbat-release-notes/39890 (дата обращения: 09.11.2026).',
    'Tanenbaum A. S., Woodhull A. S. Operating Systems: Design and Implementation. – 3rd ed. – Upper Saddle River : Prentice Hall, 2006. – 1054 p.',
    'Herder J. N., Bos H., Gras B., Homburg P., Tanenbaum A. S. MINIX 3: A Highly Reliable, Self-Repairing Operating System // ACM SIGOPS Operating Systems Review. – 2006. – Vol. 40, No. 3. – P. 80–89.',
    'MINIX 3 [Электронный ресурс] : официальный сайт проекта. – URL: https://www.minix3.org/ (дата обращения: 09.11.2026).',
    'Таненбаум Э., Бос Х. Современные операционные системы. – 4-е изд. – СПб. : Питер, 2015. – 1120 с.',
  ]),
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
      ...titlePage, ...toc, ...intro, ...goals, ...tools,
      ...ch1, ...ch2, ...ch3, ...ch4, ...ch5, ...ch6, ...ch7,
      ...conclusion, ...sources,
    ],
  }],
});

if (process.env.MARKERS) {
  fs.writeFileSync(process.env.MARKERS, JSON.stringify({ tables: tableMarkers, toc: TOC_ITEMS.map((t) => t[0]) }));
}
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log('written', OUT); });

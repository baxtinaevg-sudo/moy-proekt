const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, BorderStyle, Table, TableRow, TableCell,
  WidthType, ShadingType, TabStopType, TabStopPosition,
  Header, Footer, PageNumber, NumberFormat
} = require("docx");
const fs = require("fs");

// ============================================================
// ЦВЕТА
// ============================================================
const COLORS = {
  primary: "1F4E79",
  accent: "2E75B6",
  dark: "333333",
  gray: "666666",
  lightGray: "999999",
  bgLight: "F2F7FB",
  white: "FFFFFF",
  sectionLine: "2E75B6",
};

// ============================================================
// УТИЛИТЫ
// ============================================================
function sectionHeading(text) {
  return [
    new Paragraph({
      spacing: { before: 300, after: 80 },
      border: {
        bottom: { style: BorderStyle.SINGLE, size: 6, color: COLORS.sectionLine },
      },
      children: [
        new TextRun({
          text: text.toUpperCase(),
          bold: true,
          size: 26,
          color: COLORS.primary,
          font: "Calibri",
        }),
      ],
    }),
  ];
}

function jobTitle(company, site) {
  const children = [
    new TextRun({ text: company, bold: true, size: 24, color: COLORS.dark, font: "Calibri" }),
  ];
  if (site) {
    children.push(new TextRun({ text: `  |  ${site}`, size: 20, color: COLORS.accent, font: "Calibri" }));
  }
  return new Paragraph({ spacing: { before: 160, after: 40 }, children });
}

function roleLine(role, period) {
  return new Paragraph({
    spacing: { before: 40, after: 60 },
    children: [
      new TextRun({ text: role, bold: true, italics: true, size: 22, color: COLORS.accent, font: "Calibri" }),
      new TextRun({ text: `\n${period}`, size: 20, color: COLORS.lightGray, font: "Calibri" }),
    ],
  });
}

function bullet(text) {
  return new Paragraph({
    spacing: { before: 30, after: 30 },
    indent: { left: 360 },
    children: [
      new TextRun({ text: "•  ", size: 21, color: COLORS.accent, font: "Calibri" }),
      new TextRun({ text: text, size: 21, color: COLORS.dark, font: "Calibri" }),
    ],
  });
}

function simpleParagraph(text, opts = {}) {
  return new Paragraph({
    spacing: { before: opts.before || 60, after: opts.after || 60 },
    alignment: opts.align || AlignmentType.LEFT,
    children: [
      new TextRun({
        text,
        size: opts.size || 22,
        color: opts.color || COLORS.dark,
        bold: opts.bold || false,
        italics: opts.italics || false,
        font: "Calibri",
      }),
    ],
  });
}

function emptyLine() {
  return new Paragraph({ spacing: { before: 40, after: 40 }, children: [] });
}

// ============================================================
// РЕЗЮМЕ
// ============================================================
async function generateResume() {
  const doc = new Document({
    creator: "Бахтина Е.К.",
    title: "Резюме — Бахтина Евгения Константиновна",
    styles: {
      default: {
        document: {
          run: { font: "Calibri", size: 22, color: COLORS.dark },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            margin: { top: 720, right: 900, bottom: 720, left: 900 },
          },
        },
        children: [
          // === HEADER ===
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 80 },
            children: [
              new TextRun({
                text: "Бахтина Евгения Константиновна",
                bold: true,
                size: 36,
                color: COLORS.primary,
                font: "Calibri",
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 40 },
            children: [
              new TextRun({
                text: "Технический менеджер / Продакт-менеджер / Руководитель лаборатории",
                bold: true,
                size: 24,
                color: COLORS.accent,
                font: "Calibri",
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
            children: [
              new TextRun({ text: "📍 Жуковский, Россия  |  📱 +7 (915) 152-19-88  |  ✉️ baxtina.evg@yandex.ru", size: 20, color: COLORS.gray, font: "Calibri" }),
            ],
          }),

          // === ОБО МНЕ ===
          ...sectionHeading("Обо мне"),
          new Paragraph({
            spacing: { before: 80, after: 100 },
            children: [
              new TextRun({
                text: "Технический менеджер и продакт-менеджер с 10+ годами опыта в отрасли строительных материалов. Экспертиза охватывает полный цикл — от разработки и производства до вывода продукта на рынок. Успешный опыт руководства лабораторией, совместных R&D проектов с клиентами и формирования стратегии продуктовых линеек. Сочетаю глубокие технические знания (химия, материаловедение) с бизнес-образованием (РАНХиГС, коммерческий директор). В настоящее время обучаюсь в аспирантуре МГСУ по специальности «Материаловедение».",
                size: 21,
                color: COLORS.dark,
                font: "Calibri",
              }),
            ],
          }),

          // === ОПЫТ РАБОТЫ ===
          ...sectionHeading("Опыт работы — 10 лет 2 месяца"),

          // --- Полипласт ---
          jobTitle("ООО «Полипласт Новомосковск»", "polyplast-un.ru"),
          roleLine("Технический менеджер / Продакт-менеджер", "Март 2023 — Октябрь 2025 · 2 года 7 месяцев"),
          bullet("Разработала и вывела на рынок [X] новых продуктов в категории сухих строительных смесей и добавок"),
          bullet("Провела [X] совместных R&D тестирований с клиентами, что привело к заключению [X] контрактов"),
          bullet("Осуществляла косвенные продажи, обеспечив рост выручки продуктового направления на [Y]%"),
          bullet("Сформировала и внедрила стратегию продуктовой линейки, включающую [X] SKU"),
          bullet("Провела бенчмаркинг конкурентов, улучшены технические характеристики [X] продуктов"),
          bullet("Улучшила и стандартизировала процесс контроля качества (УТП продукта)"),
          bullet("Составила [X] технических карт и описаний продуктов для производства и коммерческого отдела"),

          // --- Качественные смеси ---
          jobTitle("ООО «Качественные смеси»", "dauer.ru"),
          roleLine("Начальник лаборатории", "Декабрь 2018 — Июнь 2021 · 2 года 7 месяцев"),
          bullet("Руководила лабораторией из [X] сотрудников, обеспечив 100% соответствие продукции ГОСТ"),
          bullet("Организовала входной контроль сырья, сократив количество брака на [Y]%"),
          bullet("Разработала [X] новых продуктов и оптимизировала [X] текущих рецептур, снизив себестоимость на [Z]%"),
          bullet("Внедрила документацию согласно НД, прошла [X] аудитов без замечаний"),
          bullet("Оптимизировала технологические процессы, сократив потери при производстве на [Y]%"),
          bullet("Составила нормы расхода сырья и рассчитала себестоимость для [X] продуктовых линеек"),
          bullet("Работала в команде маркетинга, участвовала в [X] обучающих семинарах для клиентов"),

          roleLine("Технолог (повышение до начальника лаборатории)", "Апрель 2018 — Декабрь 2018 · 9 месяцев"),
          bullet("Внедрила производственный контроль на всех этапах производства"),
          bullet("Разработала [X] новых видов продукции и улучшила [X] текущих рецептур"),
          bullet("Контролировала качество входящего сырья и готовой продукции"),
          bullet("Работала в системе 1С:Производство"),

          // --- Dauer ---
          jobTitle("ООО «Dauer»", "Жуковский, dauer.ru"),
          roleLine("Инженер-технолог", "Июнь 2017 — Март 2018 · 10 месяцев"),
          bullet("Разработала этапы технологического процесса для [X] новых продуктов"),
          bullet("Внедрила систему контроля продукции на этапах производства, повысив качество на [Y]%"),
          bullet("Получила предложение о повышении в ООО «Качественные смеси» (связанная компания)"),

          // --- ИП ---
          jobTitle("ИП Бахтина Е.К.", "Жуковский"),
          roleLine("Индивидуальный предприниматель", "Январь 2015 — Сентябрь 2017 · 2 года 9 месяцев"),
          bullet("Управляла бизнесом по оптовой и розничной торговле пищевыми продуктами"),
          bullet("Обеспечила товарооборот [X] руб./мес. с маржинальностью [Y]%"),
          bullet("Самостоятельно вела закупки, логистику, продажи и финансовый учёт"),

          // --- Аркада-Строй ---
          jobTitle("«Компания Аркада-Строй»", "Жуковский"),
          roleLine("Технолог", "Июнь 2009 — Январь 2015 · 5 лет 8 месяцев"),
          bullet("Контролировала качество и учёт сырья на производстве строительных материалов"),
          bullet("Разработала технологические процессы для [X] новых видов продукции"),
          bullet("Внедрила контроль качества на всех этапах производства"),
          bullet("Освоила систему 1С:Производство для учёта сырья и готовой продукции"),

          // --- Гарантия-Строй ---
          jobTitle("ООО «Компания Гарантия-Строй»", "Жуковский"),
          roleLine("Инженер-технолог", "Июль 2008 — Июнь 2009 · 1 год"),
          bullet("Первое место работы после окончания ВУЗа с дипломом с отличием"),
          bullet("Освоила полный цикл производственного контроля качества"),

          // === ОБРАЗОВАНИЕ ===
          ...sectionHeading("Образование"),

          simpleParagraph("ДГТУ (НПИ) — Диплом с отличием", { bold: true, size: 22, before: 100 }),
          simpleParagraph("2008 · Химический, кафедра ХВВНиСМ", { color: COLORS.gray, size: 20 }),

          simpleParagraph("МГСУ — Аспирантура (в процессе)", { bold: true, size: 22, before: 140 }),
          simpleParagraph("2024–2028 · Материаловедение · Роль: преподаватель-исследователь", { color: COLORS.gray, size: 20 }),

          simpleParagraph("ИБДА РАНХиГС — Профпереподготовка", { bold: true, size: 22, before: 140 }),
          simpleParagraph("2021–2022 · Институт бизнеса и делового администрирования · «Коммерческий директор»", { color: COLORS.gray, size: 20 }),

          // === КЛЮЧЕВЫЕ НАВЫКИ ===
          ...sectionHeading("Ключевые навыки"),

          simpleParagraph("Управление:  Руководство коллективом, управление проектами, стратегическое планирование", { before: 80 }),
          simpleParagraph("Производство:  Производственный контроль, контроль качества, технология производства строительных материалов"),
          simpleParagraph("Продукт:  Разработка новых продуктов, бенчмаркинг, формирование продуктовой линейки, технические карты"),
          simpleParagraph("Рынок:  Сухие строительные смеси, добавки для ССС, керамика, знание поставщиков и конкурентов"),
          simpleParagraph("Инструменты:  1С:Производство, R&D процессы, ГОСТ/НД"),
          simpleParagraph("Продажи:  Опыт косвенных продаж, клиентоориентированность, B2C"),

          // === ЯЗЫКИ ===
          ...sectionHeading("Языки"),
          simpleParagraph("Русский — родной    |    Английский — B1", { before: 80 }),

          // === ПРИМЕЧАНИЕ ===
          emptyLine(),
          new Paragraph({
            spacing: { before: 100 },
            border: {
              top: { style: BorderStyle.SINGLE, size: 3, color: COLORS.lightGray },
            },
            children: [
              new TextRun({
                text: "📌 Примечание: Места с [X], [Y], [Z] — плейсхолдеры для реальных цифр. Заполните их для максимального эффекта.",
                italics: true,
                size: 18,
                color: COLORS.lightGray,
                font: "Calibri",
              }),
            ],
          }),
        ],
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync("c:/Users/Нюся/moy-proekt/Резюме_Бахтина_ЕК.docx", buffer);
  console.log("✅ Резюме сохранено: Резюме_Бахтина_ЕК.docx");
}

// ============================================================
// СОПРОВОДИТЕЛЬНОЕ ПИСЬМО
// ============================================================
async function generateCoverLetter() {
  const doc = new Document({
    creator: "Бахтина Е.К.",
    title: "Сопроводительное письмо — Бахтина Евгения Константиновна",
    styles: {
      default: {
        document: {
          run: { font: "Calibri", size: 22, color: COLORS.dark },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            margin: { top: 900, right: 1000, bottom: 900, left: 1000 },
          },
        },
        children: [
          // Шапка
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            spacing: { after: 40 },
            children: [
              new TextRun({ text: "Бахтина Евгения Константиновна", bold: true, size: 22, color: COLORS.primary, font: "Calibri" }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            spacing: { after: 20 },
            children: [
              new TextRun({ text: "+7 (915) 152-19-88", size: 20, color: COLORS.gray, font: "Calibri" }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            spacing: { after: 20 },
            children: [
              new TextRun({ text: "baxtina.evg@yandex.ru", size: 20, color: COLORS.accent, font: "Calibri" }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            spacing: { after: 200 },
            children: [
              new TextRun({ text: "Жуковский, Московская область", size: 20, color: COLORS.gray, font: "Calibri" }),
            ],
          }),

          // Дата
          new Paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun({ text: "Дата: [указать дату]", size: 20, color: COLORS.gray, font: "Calibri" }),
            ],
          }),

          // Получатель
          new Paragraph({
            spacing: { after: 20 },
            children: [
              new TextRun({ text: "[Имя рекрутера или «Уважаемый руководитель!»]", size: 22, font: "Calibri" }),
            ],
          }),
          new Paragraph({
            spacing: { after: 20 },
            children: [
              new TextRun({ text: "[Название компании]", size: 22, font: "Calibri" }),
            ],
          }),
          new Paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun({ text: "[Должность / Отдел]", size: 22, font: "Calibri" }),
            ],
          }),

          // Приветствие
          new Paragraph({
            spacing: { after: 160 },
            children: [
              new TextRun({ text: "Уважаемый(ая) [Имя]!", size: 22, font: "Calibri" }),
            ],
          }),

          // Основной текст
          new Paragraph({
            spacing: { after: 120 },
            children: [
              new TextRun({
                text: "Меня заинтересовала вакансия ",
                size: 22,
                font: "Calibri",
              }),
              new TextRun({
                text: "[Название должности]",
                bold: true,
                size: 22,
                font: "Calibri",
              }),
              new TextRun({
                text: " в Вашей компании. С моим более чем 10-летним опытом работы в отрасли строительных материалов — от инженера-технолога до технического менеджера и продакт-менеджера — я уверена, что смогу стать ценным дополнением Вашей команды.",
                size: 22,
                font: "Calibri",
              }),
            ],
          }),

          new Paragraph({
            spacing: { after: 120 },
            children: [
              new TextRun({
                text: "Ключевые компетенции, которые я могу предложить:",
                size: 22,
                font: "Calibri",
              }),
            ],
          }),

          new Paragraph({
            spacing: { before: 30, after: 30 },
            indent: { left: 360 },
            children: [
              new TextRun({ text: "•  ", size: 21, color: COLORS.accent, font: "Calibri" }),
              new TextRun({
                text: "Разработка и вывод на рынок новых продуктов ",
                size: 21,
                font: "Calibri",
              }),
              new TextRun({
                text: "— опыт создания продуктовых линеек сухих строительных смесей и добавок, проведения совместных R&D проектов с ключевыми клиентами.",
                size: 21,
                color: COLORS.gray,
                font: "Calibri",
              }),
            ],
          }),
          new Paragraph({
            spacing: { before: 30, after: 30 },
            indent: { left: 360 },
            children: [
              new TextRun({ text: "•  ", size: 21, color: COLORS.accent, font: "Calibri" }),
              new TextRun({
                text: "Управление качеством и производством ",
                size: 21,
                font: "Calibri",
              }),
              new TextRun({
                text: "— руководство лабораторией, внедрение контроля качества на всех этапах, оптимизация технологических процессов и снижение себестоимости.",
                size: 21,
                color: COLORS.gray,
                font: "Calibri",
              }),
            ],
          }),
          new Paragraph({
            spacing: { before: 30, after: 30 },
            indent: { left: 360 },
            children: [
              new TextRun({ text: "•  ", size: 21, color: COLORS.accent, font: "Calibri" }),
              new TextRun({
                text: "Глубокое знание рынка ",
                size: 21,
                font: "Calibri",
              }),
              new TextRun({
                text: "— бенчмаркинг конкурентов, понимание потребностей клиентов, знание поставщиков сырья и добавок для ССС.",
                size: 21,
                color: COLORS.gray,
                font: "Calibri",
              }),
            ],
          }),
          new Paragraph({
            spacing: { before: 30, after: 30 },
            indent: { left: 360 },
            children: [
              new TextRun({ text: "•  ", size: 21, color: COLORS.accent, font: "Calibri" }),
              new TextRun({
                text: "Сочетание технической и бизнес-экспертизы ",
                size: 21,
                font: "Calibri",
              }),
              new TextRun({
                text: "— профильное химическое образование (диплом с отличием), аспирантура по материаловедению (МГСУ) и переподготовка «Коммерческий директор» (РАНХиГС).",
                size: 21,
                color: COLORS.gray,
                font: "Calibri",
              }),
            ],
          }),

          emptyLine(),

          new Paragraph({
            spacing: { after: 120 },
            children: [
              new TextRun({
                text: "В моей последней позиции в ООО «Полипласт Новомосковск» я отвечала за разработку стратегии продуктовой линейки, совместные тестирования с R&D отделами клиентов и косвенные продажи. До этого, руководя лабораторией в ООО «Качественные смеси», я обеспечила полное соответствие выпускаемой продукции требованиям ГОСТ и успешно оптимизировала себестоимость продукции.",
                size: 22,
                font: "Calibri",
              }),
            ],
          }),

          new Paragraph({
            spacing: { after: 120 },
            children: [
              new TextRun({
                text: "Я мотивирована возможностью применять свои знания и опыт для развития продуктового портфеля и укрепления позиций компании на рынке. Буду рада обсудить, как мой опыт может быть полезен [Название компании].",
                size: 22,
                font: "Calibri",
              }),
            ],
          }),

          emptyLine(),

          // Закрытие
          new Paragraph({
            spacing: { after: 20 },
            children: [
              new TextRun({ text: "С уважением,", size: 22, font: "Calibri" }),
            ],
          }),
          new Paragraph({
            spacing: { after: 20 },
            children: [
              new TextRun({ text: "Бахтина Евгения Константиновна", bold: true, size: 22, color: COLORS.primary, font: "Calibri" }),
            ],
          }),
          new Paragraph({
            spacing: { after: 20 },
            children: [
              new TextRun({ text: "+7 (915) 152-19-88", size: 20, color: COLORS.gray, font: "Calibri" }),
            ],
          }),
          new Paragraph({
            spacing: { after: 20 },
            children: [
              new TextRun({ text: "baxtina.evg@yandex.ru", size: 20, color: COLORS.accent, font: "Calibri" }),
            ],
          }),
        ],
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync("c:/Users/Нюся/moy-proekt/Сопроводительное_письмо_Бахтина_ЕК.docx", buffer);
  console.log("✅ Сопроводительное письмо сохранено: Сопроводительное_письмо_Бахтина_ЕК.docx");
}

// ============================================================
// ЗАПУСК
// ============================================================
(async () => {
  try {
    await generateResume();
    await generateCoverLetter();
    console.log("\n🎉 Все документы успешно сгенерированы!");
  } catch (err) {
    console.error("❌ Ошибка:", err);
  }
})();
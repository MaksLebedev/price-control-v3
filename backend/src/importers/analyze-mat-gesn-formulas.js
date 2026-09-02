// ============================================================
// PRICE CONTROL V3
// АНАЛИЗ ФОРМУЛ ПЕРЕСЧЁТА ЦЕН В "МАТ ГЭСН"
//
// Цель этого скрипта:
//
// понять, как существующий Excel пересчитывает:
//
// H → J
// I → K
//
// Это необходимо перед окончательным проектированием
// поля conversion_factor в PostgreSQL.
//
// Скрипт ТОЛЬКО читает Excel.
//
// Он НЕ:
// - изменяет Excel;
// - подключается к PostgreSQL;
// - изменяет структуру базы;
// - записывает какие-либо данные.
// ============================================================

const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");


// ------------------------------------------------------------
// ОСНОВНЫЕ НАСТРОЙКИ
// ------------------------------------------------------------

const TARGET_SHEET_NAME = "МАТ ГЭСН";

const excelFilePath = process.argv[2];

if (!excelFilePath) {
  console.error("Не указан путь к Excel-файлу.");
  process.exit(1);
}

const resolvedExcelPath = path.resolve(excelFilePath);

if (!fs.existsSync(resolvedExcelPath)) {
  console.error(
    `Excel-файл не найден: ${resolvedExcelPath}`
  );

  process.exit(1);
}


// ------------------------------------------------------------
// ПУТЬ К ДИАГНОСТИЧЕСКОМУ ОТЧЁТУ
// ------------------------------------------------------------

const reportsDirectory = path.resolve(
  __dirname,
  "../../reports"
);

const reportFilePath = path.join(
  reportsDirectory,
  "mat-gesn-formulas.txt"
);


// ------------------------------------------------------------
// ПРИВОДИМ ЗНАЧЕНИЕ К ТЕКСТУ
// ------------------------------------------------------------

function text(value) {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value).trim();
}


// ------------------------------------------------------------
// НОРМАЛИЗАЦИЯ ФОРМУЛЫ ДЛЯ ГРУППИРОВКИ
//
// Нам важно понять не конкретные номера строк,
// а типы используемых формул.
//
// Например:
//
// =H25/F25
// =H26/F26
//
// логически являются одной и той же схемой.
//
// Поэтому номера строк заменяем на символ #:
//
// =H#/F#
// ------------------------------------------------------------

function normalizeFormula(formula) {
  return text(formula)
    .toUpperCase()
    .replace(/\$([A-Z]+)\$\d+/g, "$1#")
    .replace(/\$([A-Z]+)\d+/g, "$1#")
    .replace(/([A-Z]+)\$\d+/g, "$1#")
    .replace(/([A-Z]+)\d+/g, "$1#")
    .replace(/\s+/g, "");
}


// ------------------------------------------------------------
// СОХРАНЯЕМ ОДИН ПРИМЕР КАЖДОГО ТИПА ФОРМУЛЫ
// ------------------------------------------------------------

function registerFormula(
  formulaStats,
  formula,
  excelRowNumber
) {
  const normalized =
    normalizeFormula(formula);

  if (!normalized) {
    return;
  }

  if (!formulaStats.has(normalized)) {
    formulaStats.set(normalized, {
      count: 0,
      examples: [],
    });
  }

  const item =
    formulaStats.get(normalized);

  item.count += 1;

  if (item.examples.length < 5) {
    item.examples.push({
      row: excelRowNumber,
      formula,
    });
  }
}


// ------------------------------------------------------------
// ОСНОВНОЙ АНАЛИЗ
// ------------------------------------------------------------

function analyzeFormulas() {
  console.log("");
  console.log(
    "Анализ формул МАТ ГЭСН..."
  );


  // ----------------------------------------------------------
  // ОТКРЫВАЕМ EXCEL
  //
  // В отличие от предыдущих анализаторов здесь нам нужны
  // непосредственно формулы ячеек, а не только их результат.
  // ----------------------------------------------------------

  const workbook = XLSX.readFile(
    resolvedExcelPath,
    {
      cellFormula: true,
      cellDates: true,
    }
  );


  if (
    !workbook.SheetNames.includes(
      TARGET_SHEET_NAME
    )
  ) {
    throw new Error(
      `Лист "${TARGET_SHEET_NAME}" не найден`
    );
  }


  const worksheet =
    workbook.Sheets[TARGET_SHEET_NAME];


  // ----------------------------------------------------------
  // ОПРЕДЕЛЯЕМ ДИАПАЗОН ЛИСТА
  // ----------------------------------------------------------

  const range = XLSX.utils.decode_range(
    worksheet["!ref"]
  );


  // ----------------------------------------------------------
  // СОБИРАЕМ СТАТИСТИКУ ОТДЕЛЬНО ДЛЯ J И K
  // ----------------------------------------------------------

  const jFormulaStats = new Map();
  const kFormulaStats = new Map();

  let jFormulaCount = 0;
  let jValueWithoutFormulaCount = 0;

  let kFormulaCount = 0;
  let kValueWithoutFormulaCount = 0;


  // ----------------------------------------------------------
  // НЕКОТОРЫЕ СТРОКИ СОХРАНИМ ПОДРОБНО
  //
  // Для них выведем:
  //
  // D/E — единицу сборника
  // F/G — единицу магазина
  // H/I — исходные цены
  // J/K — результат
  // формулы J/K
  //
  // Это поможет понять физический смысл пересчёта.
  // ----------------------------------------------------------

  const detailedExamples = [];


  // ----------------------------------------------------------
  // ПРОХОД ПО СТРОКАМ EXCEL
  //
  // Строка 1 содержит заголовки,
  // поэтому начинаем с Excel-строки 2.
  // ----------------------------------------------------------

  for (
    let rowIndex = 1;
    rowIndex <= range.e.r;
    rowIndex += 1
  ) {
    const excelRowNumber =
      rowIndex + 1;


    // --------------------------------------------------------
    // ПОЛУЧАЕМ ЯЧЕЙКИ НУЖНЫХ КОЛОНОК
    // --------------------------------------------------------

    const cellD =
      worksheet[`D${excelRowNumber}`];

    const cellE =
      worksheet[`E${excelRowNumber}`];

    const cellF =
      worksheet[`F${excelRowNumber}`];

    const cellG =
      worksheet[`G${excelRowNumber}`];

    const cellH =
      worksheet[`H${excelRowNumber}`];

    const cellI =
      worksheet[`I${excelRowNumber}`];

    const cellJ =
      worksheet[`J${excelRowNumber}`];

    const cellK =
      worksheet[`K${excelRowNumber}`];


    // --------------------------------------------------------
    // ФОРМУЛА J
    //
    // В XLSX:
    //
    // cell.f — сама формула
    // cell.v — рассчитанное значение, сохранённое в Excel
    // --------------------------------------------------------

    if (cellJ?.f) {
      jFormulaCount += 1;

      registerFormula(
        jFormulaStats,
        cellJ.f,
        excelRowNumber
      );
    } else if (
      cellJ &&
      cellJ.v !== undefined &&
      cellJ.v !== null &&
      text(cellJ.v) !== ""
    ) {
      jValueWithoutFormulaCount += 1;
    }


    // --------------------------------------------------------
    // ФОРМУЛА K
    // --------------------------------------------------------

    if (cellK?.f) {
      kFormulaCount += 1;

      registerFormula(
        kFormulaStats,
        cellK.f,
        excelRowNumber
      );
    } else if (
      cellK &&
      cellK.v !== undefined &&
      cellK.v !== null &&
      text(cellK.v) !== ""
    ) {
      kValueWithoutFormulaCount += 1;
    }


    // --------------------------------------------------------
    // СОБИРАЕМ ДО 100 ПОДРОБНЫХ ПРИМЕРОВ
    //
    // Берём только строки, где J или K действительно
    // содержат формулу.
    // --------------------------------------------------------

    if (
      detailedExamples.length < 100 &&
      (cellJ?.f || cellK?.f)
    ) {
      detailedExamples.push({
        row: excelRowNumber,

        D: cellD?.v ?? "",
        E: cellE?.v ?? "",

        F: cellF?.v ?? "",
        G: cellG?.v ?? "",

        H: cellH?.v ?? "",
        I: cellI?.v ?? "",

        J: cellJ?.v ?? "",
        JFormula: cellJ?.f ?? "",

        K: cellK?.v ?? "",
        KFormula: cellK?.f ?? "",
      });
    }
  }


  // ==========================================================
  // СОЗДАЁМ ОТЧЁТ
  // ==========================================================

  const report = [];

  report.push(
    "================================================"
  );

  report.push(
    "PRICE CONTROL V3 — ФОРМУЛЫ МАТ ГЭСН"
  );

  report.push(
    "================================================"
  );

  report.push("");

  report.push(
    `Excel: ${resolvedExcelPath}`
  );

  report.push(
    `Лист: ${TARGET_SHEET_NAME}`
  );

  report.push("");


  // ----------------------------------------------------------
  // ОБЩАЯ СТАТИСТИКА
  // ----------------------------------------------------------

  report.push(
    "1. ОБЩАЯ СТАТИСТИКА"
  );

  report.push("");

  report.push(
    `J: ячеек с формулой = ${jFormulaCount}`
  );

  report.push(
    `J: значений без формулы = ${jValueWithoutFormulaCount}`
  );

  report.push("");

  report.push(
    `K: ячеек с формулой = ${kFormulaCount}`
  );

  report.push(
    `K: значений без формулы = ${kValueWithoutFormulaCount}`
  );

  report.push("");


  // ----------------------------------------------------------
  // ТИПЫ ФОРМУЛ J
  // ----------------------------------------------------------

  report.push(
    "================================================"
  );

  report.push(
    "2. ТИПЫ ФОРМУЛ В J"
  );

  report.push(
    "================================================"
  );

  report.push("");


  [...jFormulaStats.entries()]
    .sort(
      (a, b) =>
        b[1].count - a[1].count
    )
    .forEach(
      ([normalizedFormula, info]) => {
        report.push(
          `${normalizedFormula} — ${info.count} строк`
        );

        info.examples.forEach(
          (example) => {
            report.push(
              `  Строка ${example.row}: =${example.formula}`
            );
          }
        );

        report.push("");
      }
    );


  // ----------------------------------------------------------
  // ТИПЫ ФОРМУЛ K
  // ----------------------------------------------------------

  report.push(
    "================================================"
  );

  report.push(
    "3. ТИПЫ ФОРМУЛ В K"
  );

  report.push(
    "================================================"
  );

  report.push("");


  [...kFormulaStats.entries()]
    .sort(
      (a, b) =>
        b[1].count - a[1].count
    )
    .forEach(
      ([normalizedFormula, info]) => {
        report.push(
          `${normalizedFormula} — ${info.count} строк`
        );

        info.examples.forEach(
          (example) => {
            report.push(
              `  Строка ${example.row}: =${example.formula}`
            );
          }
        );

        report.push("");
      }
    );


  // ----------------------------------------------------------
  // ПОДРОБНЫЕ ПРИМЕРЫ
  // ----------------------------------------------------------

  report.push(
    "================================================"
  );

  report.push(
    "4. ПОДРОБНЫЕ ПРИМЕРЫ"
  );

  report.push(
    "================================================"
  );

  report.push("");


  detailedExamples.forEach(
    (item) => {
      report.push(
        `Строка Excel ${item.row}`
      );

      report.push(
        `  D = ${item.D}`
      );

      report.push(
        `  E = ${item.E}`
      );

      report.push(
        `  F = ${item.F}`
      );

      report.push(
        `  G = ${item.G}`
      );

      report.push(
        `  H = ${item.H}`
      );

      report.push(
        `  I = ${item.I}`
      );

      report.push(
        `  J = ${item.J}`
      );

      report.push(
        `  Формула J = ${item.JFormula || "[нет]"}`
      );

      report.push(
        `  K = ${item.K}`
      );

      report.push(
        `  Формула K = ${item.KFormula || "[нет]"}`
      );

      report.push("");
    }
  );


  // ----------------------------------------------------------
  // ФИНАЛ
  // ----------------------------------------------------------

  report.push(
    "================================================"
  );

  report.push(
    "PostgreSQL НЕ изменялся."
  );

  report.push(
    "Excel НЕ изменялся."
  );

  report.push(
    "================================================"
  );


  // ----------------------------------------------------------
  // СОХРАНЯЕМ ОТЧЁТ
  // ----------------------------------------------------------

  fs.mkdirSync(
    reportsDirectory,
    {
      recursive: true,
    }
  );

  fs.writeFileSync(
    reportFilePath,
    report.join("\r\n"),
    "utf8"
  );


  console.log("");
  console.log(
    "Анализ формул завершён."
  );

  console.log("");

  console.log(
    "Отчёт сохранён:"
  );

  console.log(
    reportFilePath
  );

  console.log("");

  console.log(
    "PostgreSQL НЕ изменялся."
  );

  console.log(
    "Excel НЕ изменялся."
  );

  console.log("");
}


// ------------------------------------------------------------
// ЗАПУСК
// ------------------------------------------------------------

try {
  analyzeFormulas();
} catch (error) {
  console.error("");
  console.error(
    "Ошибка анализа формул:"
  );

  console.error(error);
  console.error("");

  process.exit(1);
}
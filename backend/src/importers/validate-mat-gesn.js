// ============================================================
// PRICE CONTROL V3
// ПРОВЕРКА КАЧЕСТВА ДАННЫХ ЛИСТА "МАТ ГЭСН"
//
// Этот скрипт выполняет диагностическую проверку данных
// перед будущим импортом в PostgreSQL.
//
// Он НЕ:
// - изменяет Excel;
// - подключается к PostgreSQL;
// - создаёт материалы;
// - удаляет или исправляет существующие данные.
//
// Результат работы сохраняется в:
//
// backend/reports/mat-gesn-validation.txt
//
// В отчёте будут:
// - проблемные строки;
// - дубли;
// - ошибки чисел;
// - ошибки дат;
// - домены поставщиков;
// - ссылки, которые не являются URL.
// ============================================================

const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

// ------------------------------------------------------------
// ОСНОВНЫЕ НАСТРОЙКИ
// ------------------------------------------------------------

const TARGET_SHEET_NAME = "МАТ ГЭСН";

// ------------------------------------------------------------
// ПОЛУЧАЕМ ПУТЬ К EXCEL ИЗ КОМАНДЫ ЗАПУСКА
//
// Пример:
//
// node src/importers/validate-mat-gesn.js "C:\...\file.xlsx"
// ------------------------------------------------------------

const excelFilePath = process.argv[2];

if (!excelFilePath) {
  console.error("");
  console.error("Не указан путь к Excel-файлу.");
  console.error("");

  process.exit(1);
}

const resolvedExcelPath = path.resolve(excelFilePath);

// ------------------------------------------------------------
// ПРОВЕРЯЕМ, ЧТО EXCEL-ФАЙЛ СУЩЕСТВУЕТ
// ------------------------------------------------------------

if (!fs.existsSync(resolvedExcelPath)) {
  console.error("");
  console.error("Excel-файл не найден:");
  console.error(resolvedExcelPath);
  console.error("");

  process.exit(1);
}

// ------------------------------------------------------------
// ПАПКА И ФАЙЛ ДЛЯ ДИАГНОСТИЧЕСКОГО ОТЧЁТА
//
// Папка reports будет создана автоматически,
// если её ещё нет.
// ------------------------------------------------------------

const reportsDirectory = path.resolve(__dirname, "../../reports");

const reportFilePath = path.join(reportsDirectory, "mat-gesn-validation.txt");

// ------------------------------------------------------------
// ПРЕОБРАЗОВАНИЕ ЗНАЧЕНИЯ В ТЕКСТ
//
// Нужна единая обработка пустых ячеек,
// чисел и обычного текста.
// ------------------------------------------------------------

function normalizeText(value) {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value).trim();
}

// ------------------------------------------------------------
// НОРМАЛИЗАЦИЯ ТЕКСТА ДЛЯ ПОИСКА ДУБЛЕЙ
//
// Например:
//
// "  Материал   А  "
//
// превращается в:
//
// "материал а"
//
// Исходный Excel при этом не меняется.
// ------------------------------------------------------------

function normalizeForComparison(value) {
  return normalizeText(value).replace(/\s+/g, " ").toLowerCase();
}

// ------------------------------------------------------------
// РАЗБОР ЧИСЕЛ ИЗ EXCEL
//
// В существующей базе встречаются:
//
// 1019
// 1,019.00
// 5,200,000.00
// 679.33
// 0,8
// -
//
// Поэтому обычного Number(value) недостаточно.
//
// Функция старается понять разные варианты записи.
// Если значение нельзя надёжно превратить в число,
// возвращается null.
// ------------------------------------------------------------

function parseFlexibleNumber(value) {
  let text = normalizeText(value);

  if (text === "") {
    return null;
  }

  // Явные значения, которые не являются ценой/числом.
  const invalidMarkers = new Set([
    "-",
    "—",
    "#VALUE!",
    "#N/A",
    "#REF!",
    "#DIV/0!",
    "#NAME?",
    "#NUM!",
    "#NULL!",
  ]);

  if (invalidMarkers.has(text.toUpperCase())) {
    return null;
  }

  // Удаляем обычные и неразрывные пробелы.
  text = text.replace(/\u00A0/g, "").replace(/\s/g, "");

  // ----------------------------------------------------------
  // ВАРИАНТ:
  //
  // 1,019.00
  // 5,200,000.00
  //
  // Запятые здесь являются разделителями тысяч.
  // ----------------------------------------------------------

  if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(text)) {
    const parsed = Number(text.replace(/,/g, ""));

    return Number.isFinite(parsed) ? parsed : null;
  }

  // ----------------------------------------------------------
  // ВАРИАНТ:
  //
  // 12,5
  //
  // Запятая используется как десятичный разделитель.
  // ----------------------------------------------------------

  if (/^-?\d+,\d+$/.test(text)) {
    const parsed = Number(text.replace(",", "."));

    return Number.isFinite(parsed) ? parsed : null;
  }

  // ----------------------------------------------------------
  // ОБЫЧНОЕ ЧИСЛО:
  //
  // 1
  // 2.5
  // 1000
  // ----------------------------------------------------------

  if (/^-?\d+(\.\d+)?$/.test(text)) {
    const parsed = Number(text);

    return Number.isFinite(parsed) ? parsed : null;
  }

  return null;
}

// ------------------------------------------------------------
// ПРОВЕРКА ДАТЫ
//
// В текущем Excel встречаются даты вида:
//
// 7/7/26
// 5/19/25
// 01.09.2026
//
// Пока задача только определить,
// похоже ли значение на корректную календарную дату.
//
// Сам формат хранения в PostgreSQL решим позднее.
// ------------------------------------------------------------

function isValidDateValue(value) {
  const text = normalizeText(value);

  if (text === "") {
    return true;
  }

  // ----------------------------------------------------------
  // Формат M/D/YY или M/D/YYYY
  // ----------------------------------------------------------

  let match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);

  if (match) {
    const month = Number(match[1]);
    const day = Number(match[2]);

    let year = Number(match[3]);

    if (year < 100) {
      year += 2000;
    }

    return isRealCalendarDate(year, month, day);
  }

  // ----------------------------------------------------------
  // Формат DD.MM.YYYY
  // ----------------------------------------------------------

  match = text.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);

  if (match) {
    const day = Number(match[1]);
    const month = Number(match[2]);
    const year = Number(match[3]);

    return isRealCalendarDate(year, month, day);
  }

  // Если формат неизвестен,
  // считаем значение подозрительным.
  return false;
}

// ------------------------------------------------------------
// ПРОВЕРКА РЕАЛЬНОСТИ КАЛЕНДАРНОЙ ДАТЫ
//
// Например:
//
// 31.02.2026
//
// не должна считаться корректной.
// ------------------------------------------------------------

function isRealCalendarDate(year, month, day) {
  if (
    !Number.isInteger(year) ||
    !Number.isInteger(month) ||
    !Number.isInteger(day)
  ) {
    return false;
  }

  if (month < 1 || month > 12 || day < 1 || day > 31) {
    return false;
  }

  const date = new Date(Date.UTC(year, month - 1, day));

  return (
    date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day
  );
}

// ------------------------------------------------------------
// ПРОВЕРКА URL
//
// Нас интересуют обычные HTTP/HTTPS-ссылки.
//
// Значения вроде:
//
// "средняя цена"
//
// URL не являются.
// ------------------------------------------------------------

function getUrlInfo(value) {
  const text = normalizeText(value);

  if (text === "") {
    return {
      isEmpty: true,
      isUrl: false,
      domain: null,
    };
  }

  try {
    const url = new URL(text);

    if (url.protocol !== "http:" && url.protocol !== "https:") {
      return {
        isEmpty: false,
        isUrl: false,
        domain: null,
      };
    }

    const domain = url.hostname.toLowerCase().replace(/^www\./, "");

    return {
      isEmpty: false,
      isUrl: true,
      domain,
    };
  } catch {
    return {
      isEmpty: false,
      isUrl: false,
      domain: null,
    };
  }
}

// ------------------------------------------------------------
// ДОБАВЛЕНИЕ СТРОКИ В СПИСОК ПРОБЛЕМ
//
// Чтобы отчёт был удобным,
// сохраняем номер строки Excel и краткое описание.
// ------------------------------------------------------------

function addIssue(issueList, excelRowNumber, description) {
  issueList.push({
    excelRowNumber,
    description,
  });
}

// ------------------------------------------------------------
// ФОРМАТИРОВАНИЕ СПИСКА ПРОБЛЕМ ДЛЯ ОТЧЁТА
//
// Чтобы отчёт не становился огромным,
// выводим максимум maxExamples конкретных строк.
//
// При этом общее количество ошибок показываем полностью.
// ------------------------------------------------------------

function appendIssuesToReport(report, title, issues, maxExamples = 50) {
  report.push("");
  report.push("================================================");
  report.push(title);
  report.push("================================================");
  report.push("");

  report.push(`Всего найдено: ${issues.length}`);

  report.push("");

  if (issues.length === 0) {
    report.push("Проблем не найдено.");
    return;
  }

  const examples = issues.slice(0, maxExamples);

  for (const issue of examples) {
    report.push(`Строка Excel ${issue.excelRowNumber}: ${issue.description}`);
  }

  if (issues.length > maxExamples) {
    report.push("");
    report.push(`Показаны первые ${maxExamples} из ${issues.length}.`);
  }
}

// ------------------------------------------------------------
// ОСНОВНАЯ ФУНКЦИЯ ПРОВЕРКИ
// ------------------------------------------------------------

function validateMatGesn() {
  console.log("");
  console.log("Проверка качества данных МАТ ГЭСН...");

  // ----------------------------------------------------------
  // ЧИТАЕМ EXCEL
  //
  // Записи обратно в книгу здесь нет.
  // ----------------------------------------------------------

  const workbook = XLSX.readFile(resolvedExcelPath, {
    cellDates: true,
    cellFormula: true,
  });

  // ----------------------------------------------------------
  // ПРОВЕРЯЕМ НАЛИЧИЕ ЛИСТА
  // ----------------------------------------------------------

  if (!workbook.SheetNames.includes(TARGET_SHEET_NAME)) {
    throw new Error(`Лист "${TARGET_SHEET_NAME}" не найден`);
  }

  const worksheet = workbook.Sheets[TARGET_SHEET_NAME];

  // ----------------------------------------------------------
  // ЧИТАЕМ ЛИСТ КАК МАССИВ СТРОК
  //
  // raw: false нужен, чтобы анализировать значения
  // примерно в том виде, как они видны пользователю в Excel.
  // ----------------------------------------------------------

  const rows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    raw: false,
    defval: "",
    blankrows: true,
  });

  if (rows.length === 0) {
    throw new Error(`Лист "${TARGET_SHEET_NAME}" пуст`);
  }

  // ----------------------------------------------------------
  // Первая строка содержит заголовки.
  // Остальные — данные.
  // ----------------------------------------------------------

  const dataRows = rows.slice(1);

  // ==========================================================
  // СПИСКИ ПРОБЛЕМ
  // ==========================================================

  const codeWithoutName = [];

  const invalidBaseQuantity = [];

  const quantityWithoutUnit = [];

  const unitWithoutQuantity = [];

  const invalidPrices = {
    H: [],
    I: [],
    J: [],
    K: [],
  };

  const invalidDates = {
    N: [],
    O: [],
  };

  const invalidUrls = {
    L: [],
    T: [],
    W: [],
  };

  // ==========================================================
  // СТАТИСТИКА ЧИСЕЛ
  // ==========================================================

  const priceStats = {
    H: {
      filled: 0,
      numeric: 0,
      invalid: 0,
    },

    I: {
      filled: 0,
      numeric: 0,
      invalid: 0,
    },

    J: {
      filled: 0,
      numeric: 0,
      invalid: 0,
    },

    K: {
      filled: 0,
      numeric: 0,
      invalid: 0,
    },
  };

  // ==========================================================
  // СТАТИСТИКА ДОМЕНОВ
  //
  // Считаем отдельно основной URL,
  // поставщика №2 и поставщика №3.
  // ==========================================================

  const domainStats = {
    L: new Map(),
    T: new Map(),
    W: new Map(),
  };

  // ==========================================================
  // ДАННЫЕ ДЛЯ ПОИСКА ДУБЛЕЙ
  // ==========================================================

  const rowsByCode = new Map();

  const rowsByCodeAndName = new Map();

  // ==========================================================
  // ПРОХОД ПО ВСЕМ СТРОКАМ "МАТ ГЭСН"
  // ==========================================================

  dataRows.forEach((row, dataIndex) => {
    // Номер строки Excel:
    //
    // dataIndex = 0 соответствует строке Excel 2.
    const excelRowNumber = dataIndex + 2;

    // ------------------------------------------------------
    // ЧИТАЕМ НУЖНЫЕ КОЛОНКИ
    // ------------------------------------------------------

    const code = normalizeText(row[0]); // A

    const officialName = normalizeText(row[1]); // B

    const baseQuantity = normalizeText(row[3]); // D

    const baseUnit = normalizeText(row[4]); // E

    // ------------------------------------------------------
    // ПОЛНОСТЬЮ ПУСТЫЕ СТРОКИ ПРОПУСКАЕМ
    //
    // Они не являются материалами.
    // ------------------------------------------------------

    const hasAnyData = row.some((value) => normalizeText(value) !== "");

    if (!hasAnyData) {
      return;
    }

    // ======================================================
    // 1. КОД ЕСТЬ, А НАИМЕНОВАНИЕ ОТСУТСТВУЕТ
    // ======================================================

    if (code !== "" && officialName === "") {
      addIssue(
        codeWithoutName,
        excelRowNumber,
        `код "${code}", наименование отсутствует`,
      );
    }

    // ======================================================
    // 2. ПРОВЕРКА КОЛИЧЕСТВА СБОРНИКА — D
    //
    // Если D заполнено,
    // ожидаем получить обычное положительное число.
    // ======================================================

    if (baseQuantity !== "") {
      const parsedQuantity = parseFlexibleNumber(baseQuantity);

      if (parsedQuantity === null || parsedQuantity <= 0) {
        addIssue(
          invalidBaseQuantity,
          excelRowNumber,
          `D="${baseQuantity}", код="${code}"`,
        );
      }
    }

    // ======================================================
    // 3. КОЛИЧЕСТВО ЕСТЬ, ЕДИНИЦЫ НЕТ
    // ======================================================

    if (baseQuantity !== "" && baseUnit === "") {
      addIssue(
        quantityWithoutUnit,
        excelRowNumber,
        `D="${baseQuantity}", но E пусто, код="${code}"`,
      );
    }

    // ======================================================
    // 4. ЕДИНИЦА ЕСТЬ, КОЛИЧЕСТВА НЕТ
    // ======================================================

    if (baseQuantity === "" && baseUnit !== "") {
      addIssue(
        unitWithoutQuantity,
        excelRowNumber,
        `E="${baseUnit}", но D пусто, код="${code}"`,
      );
    }

    // ======================================================
    // 5. СОБИРАЕМ ДАННЫЕ ДЛЯ ПОИСКА ДУБЛЕЙ
    // ======================================================

    if (code !== "") {
      const normalizedCode = normalizeForComparison(code);

      if (!rowsByCode.has(normalizedCode)) {
        rowsByCode.set(normalizedCode, []);
      }

      rowsByCode.get(normalizedCode).push(excelRowNumber);

      const normalizedName = normalizeForComparison(officialName);

      const combinedKey = `${normalizedCode}||${normalizedName}`;

      if (!rowsByCodeAndName.has(combinedKey)) {
        rowsByCodeAndName.set(combinedKey, []);
      }

      rowsByCodeAndName.get(combinedKey).push(excelRowNumber);
    }

    // ======================================================
    // 6. ПРОВЕРКА ЦЕН H / I / J / K
    // ======================================================

    const priceColumns = {
      H: row[7],
      I: row[8],
      J: row[9],
      K: row[10],
    };

    for (const [columnName, rawValue] of Object.entries(priceColumns)) {
      const value = normalizeText(rawValue);

      if (value === "") {
        continue;
      }

      priceStats[columnName].filled += 1;

      const parsed = parseFlexibleNumber(value);

      if (parsed !== null) {
        priceStats[columnName].numeric += 1;
      } else {
        priceStats[columnName].invalid += 1;

        addIssue(
          invalidPrices[columnName],
          excelRowNumber,
          `${columnName}="${value}", код="${code}"`,
        );
      }
    }

    // ======================================================
    // 7. ПРОВЕРКА ДАТ N / O
    // ======================================================

    const dateColumns = {
      N: row[13],
      O: row[14],
    };

    for (const [columnName, rawValue] of Object.entries(dateColumns)) {
      const value = normalizeText(rawValue);

      if (value === "") {
        continue;
      }

      if (!isValidDateValue(value)) {
        addIssue(
          invalidDates[columnName],
          excelRowNumber,
          `${columnName}="${value}", код="${code}"`,
        );
      }
    }

    // ======================================================
    // 8. ПРОВЕРКА ССЫЛОК L / T / W
    //
    // Одновременно собираем домены поставщиков.
    // ======================================================

    const urlColumns = {
      L: row[11],
      T: row[19],
      W: row[22],
    };

    for (const [columnName, rawValue] of Object.entries(urlColumns)) {
      const value = normalizeText(rawValue);

      if (value === "") {
        continue;
      }

      const urlInfo = getUrlInfo(value);

      if (!urlInfo.isUrl) {
        addIssue(
          invalidUrls[columnName],
          excelRowNumber,
          `${columnName}="${value}", код="${code}"`,
        );

        continue;
      }

      // ----------------------------------------------------
      // Увеличиваем счётчик найденного домена.
      // ----------------------------------------------------

      const currentCount = domainStats[columnName].get(urlInfo.domain) || 0;

      domainStats[columnName].set(urlInfo.domain, currentCount + 1);
    }
  });

  // ==========================================================
  // ФОРМИРУЕМ СПИСКИ ДУБЛЕЙ
  // ==========================================================

  const duplicateCodes = [];

  for (const [code, excelRows] of rowsByCode.entries()) {
    if (excelRows.length <= 1) {
      continue;
    }

    duplicateCodes.push({
      code,
      count: excelRows.length,
      excelRows,
    });
  }

  const duplicateCodeAndName = [];

  for (const [key, excelRows] of rowsByCodeAndName.entries()) {
    if (excelRows.length <= 1) {
      continue;
    }

    const separatorIndex = key.indexOf("||");

    const code = key.slice(0, separatorIndex);

    const name = key.slice(separatorIndex + 2);

    duplicateCodeAndName.push({
      code,
      name,
      count: excelRows.length,
      excelRows,
    });
  }

  // ==========================================================
  // СОЗДАЁМ ТЕКСТОВЫЙ ОТЧЁТ
  // ==========================================================

  const report = [];

  report.push("================================================");

  report.push("PRICE CONTROL V3 — ПРОВЕРКА МАТ ГЭСН");

  report.push("================================================");

  report.push("");

  report.push(`Excel: ${resolvedExcelPath}`);

  report.push(`Лист: ${TARGET_SHEET_NAME}`);

  report.push(`Всего строк данных: ${dataRows.length}`);

  report.push("");

  report.push("PostgreSQL при проверке НЕ изменялся.");

  report.push("Excel при проверке НЕ изменялся.");

  // ==========================================================
  // ОБЩИЕ ПРОБЛЕМЫ СТРУКТУРЫ
  // ==========================================================

  appendIssuesToReport(
    report,
    "1. Код ресурса есть, но наименование отсутствует",
    codeWithoutName,
  );

  appendIssuesToReport(
    report,
    "2. Некорректное количество единиц сборника — колонка D",
    invalidBaseQuantity,
  );

  appendIssuesToReport(
    report,
    "3. Количество D заполнено, но единица E отсутствует",
    quantityWithoutUnit,
  );

  appendIssuesToReport(
    report,
    "4. Единица E заполнена, но количество D отсутствует",
    unitWithoutQuantity,
  );

  // ==========================================================
  // ДУБЛИ КОДОВ
  // ==========================================================

  report.push("");
  report.push("================================================");

  report.push("5. Дубли resource_code");

  report.push("================================================");

  report.push("");

  report.push(
    `Количество кодов, встречающихся более одного раза: ${duplicateCodes.length}`,
  );

  report.push("");

  duplicateCodes.slice(0, 100).forEach((item) => {
    report.push(
      `Код "${item.code}" — ${item.count} строк: ${item.excelRows.join(", ")}`,
    );
  });

  if (duplicateCodes.length > 100) {
    report.push("");
    report.push(`Показаны первые 100 из ${duplicateCodes.length}.`);
  }

  // ==========================================================
  // ДУБЛИ КОД + НАИМЕНОВАНИЕ
  // ==========================================================

  report.push("");
  report.push("================================================");

  report.push("6. Дубли resource_code + official_name");

  report.push("================================================");

  report.push("");

  report.push(
    `Количество повторяющихся комбинаций: ${duplicateCodeAndName.length}`,
  );

  report.push("");

  duplicateCodeAndName.slice(0, 100).forEach((item) => {
    report.push(
      `Код "${item.code}", строк ${item.count}: ${item.excelRows.join(", ")}`,
    );

    report.push(`  Наименование: ${item.name}`);
  });

  if (duplicateCodeAndName.length > 100) {
    report.push("");
    report.push(`Показаны первые 100 из ${duplicateCodeAndName.length}.`);
  }

  // ==========================================================
  // СТАТИСТИКА ЦЕН
  // ==========================================================

  report.push("");
  report.push("================================================");

  report.push("7. Проверка цен H / I / J / K");

  report.push("================================================");

  report.push("");

  for (const columnName of ["H", "I", "J", "K"]) {
    const stats = priceStats[columnName];

    report.push(`Колонка ${columnName}:`);

    report.push(`  заполнено: ${stats.filled}`);

    report.push(`  распознано как число: ${stats.numeric}`);

    report.push(`  не удалось распознать: ${stats.invalid}`);

    report.push("");
  }

  for (const columnName of ["H", "I", "J", "K"]) {
    appendIssuesToReport(
      report,
      `7.${columnName}. Нераспознанные значения цены ${columnName}`,
      invalidPrices[columnName],
    );
  }

  // ==========================================================
  // ДАТЫ
  // ==========================================================

  appendIssuesToReport(
    report,
    "8.N. Некорректные даты парсинга — колонка N",
    invalidDates.N,
  );

  appendIssuesToReport(
    report,
    "8.O. Некорректные ручные даты — колонка O",
    invalidDates.O,
  );

  // ==========================================================
  // НЕПРАВИЛЬНЫЕ URL
  // ==========================================================

  appendIssuesToReport(
    report,
    "9.L. Значения основной ссылки L, которые не являются URL",
    invalidUrls.L,
  );

  appendIssuesToReport(
    report,
    "9.T. Значения ссылки поставщика 2 — T, которые не являются URL",
    invalidUrls.T,
  );

  appendIssuesToReport(
    report,
    "9.W. Значения ссылки поставщика 3 — W, которые не являются URL",
    invalidUrls.W,
  );

  // ==========================================================
  // ДОМЕНЫ ПОСТАВЩИКОВ
  // ==========================================================

  report.push("");
  report.push("================================================");

  report.push("10. Домены поставщиков из L / T / W");

  report.push("================================================");

  report.push("");

  for (const columnName of ["L", "T", "W"]) {
    report.push(`Колонка ${columnName}:`);

    const sortedDomains = [...domainStats[columnName].entries()].sort(
      (a, b) => b[1] - a[1],
    );

    if (sortedDomains.length === 0) {
      report.push("  корректных URL не найдено");
    } else {
      for (const [domain, count] of sortedDomains) {
        report.push(`  ${domain}: ${count}`);
      }
    }

    report.push("");
  }

  // ==========================================================
  // ИТОГОВАЯ СВОДКА
  // ==========================================================

  report.push("");
  report.push("================================================");

  report.push("11. ИТОГОВАЯ СВОДКА");

  report.push("================================================");

  report.push("");

  report.push(`Код есть, название отсутствует: ${codeWithoutName.length}`);

  report.push(`Некорректное D: ${invalidBaseQuantity.length}`);

  report.push(`D есть, E нет: ${quantityWithoutUnit.length}`);

  report.push(`E есть, D нет: ${unitWithoutQuantity.length}`);

  report.push(`Дубли кодов: ${duplicateCodes.length}`);

  report.push(`Дубли код + название: ${duplicateCodeAndName.length}`);

  report.push(`Некорректные даты N: ${invalidDates.N.length}`);

  report.push(`Некорректные даты O: ${invalidDates.O.length}`);

  report.push(`Не-URL в L: ${invalidUrls.L.length}`);

  report.push(`Не-URL в T: ${invalidUrls.T.length}`);

  report.push(`Не-URL в W: ${invalidUrls.W.length}`);

  report.push("");

  report.push("Проверка завершена.");

  report.push("PostgreSQL НЕ изменялся.");

  report.push("Excel НЕ изменялся.");

  report.push("");

  // ==========================================================
  // СОХРАНЯЕМ ОТЧЁТ НА ДИСК
  // ==========================================================

  fs.mkdirSync(reportsDirectory, {
    recursive: true,
  });

  fs.writeFileSync(reportFilePath, report.join("\r\n"), "utf8");

  // ----------------------------------------------------------
  // В ТЕРМИНАЛ ВЫВОДИМ ТОЛЬКО КРАТКИЙ РЕЗУЛЬТАТ
  //
  // Полный анализ находится в текстовом отчёте.
  // ----------------------------------------------------------

  console.log("");
  console.log("Проверка завершена.");

  console.log("");

  console.log("Отчёт сохранён:");

  console.log(reportFilePath);

  console.log("");

  console.log("PostgreSQL НЕ изменялся.");

  console.log("Excel НЕ изменялся.");

  console.log("");
}

// ------------------------------------------------------------
// ЗАПУСК СКРИПТА
//
// Если возникает ошибка,
// она выводится в терминал,
// а никакие данные при этом не изменяются.
// ------------------------------------------------------------

try {
  validateMatGesn();
} catch (error) {
  console.error("");
  console.error("Ошибка проверки МАТ ГЭСН:");

  console.error(error);
  console.error("");

  process.exit(1);
}

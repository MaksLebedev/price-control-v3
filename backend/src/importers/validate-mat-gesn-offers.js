// ============================================================
// PRICE CONTROL V3
// ПРОВЕРКА СВЯЗНОСТИ ДАННЫХ ПОСТАВЩИКОВ В "МАТ ГЭСН"
//
// Скрипт нужен перед переносом supplier_offers в PostgreSQL.
//
// Он ТОЛЬКО читает Excel и создаёт диагностический отчёт.
//
// Он НЕ:
// - изменяет Excel;
// - подключается к PostgreSQL;
// - добавляет поставщиков;
// - добавляет предложения;
// - исправляет исходные значения.
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
  console.error(`Excel-файл не найден: ${resolvedExcelPath}`);
  process.exit(1);
}

// ------------------------------------------------------------
// ПУТЬ К ОТЧЁТУ
// ------------------------------------------------------------

const reportsDirectory = path.resolve(__dirname, "../../reports");

const reportFilePath = path.join(
  reportsDirectory,
  "mat-gesn-offers-validation.txt",
);

// ------------------------------------------------------------
// ПРИВОДИМ ЛЮБОЕ ЗНАЧЕНИЕ EXCEL К ТЕКСТУ
// ------------------------------------------------------------

function text(value) {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value).trim();
}

// ------------------------------------------------------------
// ПРОВЕРЯЕМ, ЯВЛЯЕТСЯ ЛИ ЗНАЧЕНИЕ HTTP/HTTPS-ССЫЛКОЙ
// ------------------------------------------------------------

function isUrl(value) {
  const valueText = text(value);

  if (!valueText) {
    return false;
  }

  try {
    const url = new URL(valueText);

    return url.protocol === "http:" || url.protocol === "https:";
  } catch {
    return false;
  }
}

// ------------------------------------------------------------
// ДОБАВЛЯЕМ ПРОБЛЕМУ В ДИАГНОСТИЧЕСКИЙ СПИСОК
// ------------------------------------------------------------

function addIssue(list, rowNumber, code, description) {
  list.push({
    rowNumber,
    code,
    description,
  });
}

// ------------------------------------------------------------
// ДОБАВЛЯЕМ РАЗДЕЛ С ПРОБЛЕМАМИ В ИТОГОВЫЙ ОТЧЁТ
//
// Чтобы файл не стал чрезмерно большим,
// показываем максимум 100 конкретных строк каждого типа.
// ------------------------------------------------------------

function addIssueSection(report, title, issues) {
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

  issues.slice(0, 100).forEach((issue) => {
    report.push(
      `Строка Excel ${issue.rowNumber}, код "${issue.code}": ${issue.description}`,
    );
  });

  if (issues.length > 100) {
    report.push("");
    report.push(`Показаны первые 100 из ${issues.length}.`);
  }
}

// ------------------------------------------------------------
// ОСНОВНАЯ ПРОВЕРКА
// ------------------------------------------------------------

function validateOffers() {
  console.log("");
  console.log("Проверка предложений поставщиков МАТ ГЭСН...");

  // ----------------------------------------------------------
  // ЧИТАЕМ EXCEL
  // ----------------------------------------------------------

  const workbook = XLSX.readFile(resolvedExcelPath, {
    cellDates: true,
    cellFormula: true,
  });

  if (!workbook.SheetNames.includes(TARGET_SHEET_NAME)) {
    throw new Error(`Лист "${TARGET_SHEET_NAME}" не найден`);
  }

  const worksheet = workbook.Sheets[TARGET_SHEET_NAME];

  const rows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    raw: false,
    defval: "",
    blankrows: true,
  });

  const dataRows = rows.slice(1);

  // ==========================================================
  // СПИСКИ ПРОБЛЕМ ОСНОВНОГО ПРЕДЛОЖЕНИЯ
  // ==========================================================

  const mainUrlWithoutQuantity = [];
  const mainQuantityWithoutUrl = [];
  const mainQuantityWithoutUnit = [];
  const mainUnitWithoutQuantity = [];

  const parsingEnabledWithoutUrl = [];
  const parsedPriceWithoutUrl = [];

  const mainNonUrlReference = [];

  // ==========================================================
  // СПИСКИ ПРОБЛЕМ ПОСТАВЩИКА №2
  // ==========================================================

  const supplier2UrlWithoutQuantity = [];
  const supplier2QuantityWithoutUrl = [];
  const supplier2QuantityWithoutUnit = [];
  const supplier2UnitWithoutQuantity = [];

  // ==========================================================
  // СПИСКИ ПРОБЛЕМ ПОСТАВЩИКА №3
  // ==========================================================

  const supplier3UrlWithoutQuantity = [];
  const supplier3QuantityWithoutUrl = [];
  const supplier3QuantityWithoutUnit = [];
  const supplier3UnitWithoutQuantity = [];

  // ==========================================================
  // ОБЩИЕ СЧЁТЧИКИ
  // ==========================================================

  const stats = {
    mainUrl: 0,
    mainTextReference: 0,

    parsingEnabled: 0,
    parsedStorePrice: 0,
    parsedCollectionPrice: 0,

    supplier2Url: 0,
    supplier3Url: 0,
  };

  // ==========================================================
  // ПРОХОДИМ ПО ВСЕМ СТРОКАМ
  // ==========================================================

  dataRows.forEach((row, index) => {
    const rowNumber = index + 2;

    const code = text(row[0]); // A

    // --------------------------------------------------------
    // ОСНОВНОЙ ПОСТАВЩИК
    //
    // F — количество магазина
    // G — единица магазина
    // I — цена магазина от парсера
    // K — цена сборника от парсера
    // L — ссылка / текстовая пометка
    // M — учтен для парсинга
    // --------------------------------------------------------

    const mainQuantity = text(row[5]); // F
    const mainUnit = text(row[6]); // G

    const parsedStorePrice = text(row[8]); // I
    const parsedCollectionPrice = text(row[10]); // K

    const mainReference = text(row[11]); // L
    const parsingFlag = text(row[12]); // M

    const mainReferenceIsUrl = isUrl(mainReference);

    if (mainReferenceIsUrl) {
      stats.mainUrl += 1;
    } else if (mainReference) {
      stats.mainTextReference += 1;

      addIssue(mainNonUrlReference, rowNumber, code, `L="${mainReference}"`);
    }

    if (mainReferenceIsUrl && !mainQuantity) {
      addIssue(
        mainUrlWithoutQuantity,
        rowNumber,
        code,
        "есть URL L, но количество F отсутствует",
      );
    }

    if (mainQuantity && !mainReference) {
      addIssue(
        mainQuantityWithoutUrl,
        rowNumber,
        code,
        `F="${mainQuantity}", но L пусто`,
      );
    }

    if (mainQuantity && !mainUnit) {
      addIssue(
        mainQuantityWithoutUnit,
        rowNumber,
        code,
        `F="${mainQuantity}", но G пусто`,
      );
    }

    if (!mainQuantity && mainUnit) {
      addIssue(
        mainUnitWithoutQuantity,
        rowNumber,
        code,
        `G="${mainUnit}", но F пусто`,
      );
    }

    // --------------------------------------------------------
    // УЧАСТИЕ В ПАРСИНГЕ
    //
    // Сейчас ожидаем, что M="да" используется только
    // вместе с настоящей URL-ссылкой.
    // --------------------------------------------------------

    if (parsingFlag) {
      stats.parsingEnabled += 1;

      if (!mainReferenceIsUrl) {
        addIssue(
          parsingEnabledWithoutUrl,
          rowNumber,
          code,
          `M="${parsingFlag}", но L не является URL`,
        );
      }
    }

    // --------------------------------------------------------
    // ПАРСИНГОВЫЕ ЦЕНЫ
    //
    // Если I или K заполнены, но URL отсутствует,
    // это нужно отдельно понять до миграции.
    // --------------------------------------------------------

    if (parsedStorePrice) {
      stats.parsedStorePrice += 1;
    }

    if (parsedCollectionPrice) {
      stats.parsedCollectionPrice += 1;
    }

    if ((parsedStorePrice || parsedCollectionPrice) && !mainReferenceIsUrl) {
      addIssue(
        parsedPriceWithoutUrl,
        rowNumber,
        code,
        `I="${parsedStorePrice}", K="${parsedCollectionPrice}", L="${mainReference}"`,
      );
    }

    // --------------------------------------------------------
    // ПОСТАВЩИК №2
    //
    // R — количество
    // S — единица
    // T — URL
    // --------------------------------------------------------

    const supplier2Quantity = text(row[17]); // R
    const supplier2Unit = text(row[18]); // S
    const supplier2Url = text(row[19]); // T

    if (supplier2Url) {
      stats.supplier2Url += 1;
    }

    if (supplier2Url && !supplier2Quantity) {
      addIssue(
        supplier2UrlWithoutQuantity,
        rowNumber,
        code,
        "есть URL T, но количество R отсутствует",
      );
    }

    if (supplier2Quantity && !supplier2Url) {
      addIssue(
        supplier2QuantityWithoutUrl,
        rowNumber,
        code,
        `R="${supplier2Quantity}", но T пусто`,
      );
    }

    if (supplier2Quantity && !supplier2Unit) {
      addIssue(
        supplier2QuantityWithoutUnit,
        rowNumber,
        code,
        `R="${supplier2Quantity}", но S пусто`,
      );
    }

    if (!supplier2Quantity && supplier2Unit) {
      addIssue(
        supplier2UnitWithoutQuantity,
        rowNumber,
        code,
        `S="${supplier2Unit}", но R пусто`,
      );
    }

    // --------------------------------------------------------
    // ПОСТАВЩИК №3
    //
    // U — количество
    // V — единица
    // W — URL
    // --------------------------------------------------------

    const supplier3Quantity = text(row[20]); // U
    const supplier3Unit = text(row[21]); // V
    const supplier3Url = text(row[22]); // W

    if (supplier3Url) {
      stats.supplier3Url += 1;
    }

    if (supplier3Url && !supplier3Quantity) {
      addIssue(
        supplier3UrlWithoutQuantity,
        rowNumber,
        code,
        "есть URL W, но количество U отсутствует",
      );
    }

    if (supplier3Quantity && !supplier3Url) {
      addIssue(
        supplier3QuantityWithoutUrl,
        rowNumber,
        code,
        `U="${supplier3Quantity}", но W пусто`,
      );
    }

    if (supplier3Quantity && !supplier3Unit) {
      addIssue(
        supplier3QuantityWithoutUnit,
        rowNumber,
        code,
        `U="${supplier3Quantity}", но V пусто`,
      );
    }

    if (!supplier3Quantity && supplier3Unit) {
      addIssue(
        supplier3UnitWithoutQuantity,
        rowNumber,
        code,
        `V="${supplier3Unit}", но U пусто`,
      );
    }
  });

  // ==========================================================
  // СОЗДАЁМ ОТЧЁТ
  // ==========================================================

  const report = [];

  report.push("================================================");

  report.push("PRICE CONTROL V3 — ПРОВЕРКА SUPPLIER OFFERS");

  report.push("================================================");

  report.push("");

  report.push(`Excel: ${resolvedExcelPath}`);
  report.push(`Лист: ${TARGET_SHEET_NAME}`);
  report.push(`Строк данных: ${dataRows.length}`);

  report.push("");

  report.push(`Основных настоящих URL L: ${stats.mainUrl}`);

  report.push(`Текстовых пометок вместо URL L: ${stats.mainTextReference}`);

  report.push(`Строк с признаком парсинга M: ${stats.parsingEnabled}`);

  report.push(`Заполненных I: ${stats.parsedStorePrice}`);

  report.push(`Заполненных K: ${stats.parsedCollectionPrice}`);

  report.push(`URL поставщика 2 — T: ${stats.supplier2Url}`);

  report.push(`URL поставщика 3 — W: ${stats.supplier3Url}`);

  // ----------------------------------------------------------
  // ОСНОВНОЕ ПРЕДЛОЖЕНИЕ
  // ----------------------------------------------------------

  addIssueSection(
    report,
    "1. Основной URL L есть, но количество F отсутствует",
    mainUrlWithoutQuantity,
  );

  addIssueSection(
    report,
    "2. Количество F есть, но L пусто",
    mainQuantityWithoutUrl,
  );

  addIssueSection(
    report,
    "3. Количество F есть, но единица G отсутствует",
    mainQuantityWithoutUnit,
  );

  addIssueSection(
    report,
    "4. Единица G есть, но количество F отсутствует",
    mainUnitWithoutQuantity,
  );

  addIssueSection(
    report,
    "5. M заполнено, но L не является URL",
    parsingEnabledWithoutUrl,
  );

  addIssueSection(
    report,
    "6. Парсинговая цена I/K есть, но L не является URL",
    parsedPriceWithoutUrl,
  );

  addIssueSection(
    report,
    "7. Текстовые пометки в L вместо URL",
    mainNonUrlReference,
  );

  // ----------------------------------------------------------
  // ПОСТАВЩИК №2
  // ----------------------------------------------------------

  addIssueSection(
    report,
    "8. Поставщик 2: T есть, но R отсутствует",
    supplier2UrlWithoutQuantity,
  );

  addIssueSection(
    report,
    "9. Поставщик 2: R есть, но T отсутствует",
    supplier2QuantityWithoutUrl,
  );

  addIssueSection(
    report,
    "10. Поставщик 2: R есть, но S отсутствует",
    supplier2QuantityWithoutUnit,
  );

  addIssueSection(
    report,
    "11. Поставщик 2: S есть, но R отсутствует",
    supplier2UnitWithoutQuantity,
  );

  // ----------------------------------------------------------
  // ПОСТАВЩИК №3
  // ----------------------------------------------------------

  addIssueSection(
    report,
    "12. Поставщик 3: W есть, но U отсутствует",
    supplier3UrlWithoutQuantity,
  );

  addIssueSection(
    report,
    "13. Поставщик 3: U есть, но W отсутствует",
    supplier3QuantityWithoutUrl,
  );

  addIssueSection(
    report,
    "14. Поставщик 3: U есть, но V отсутствует",
    supplier3QuantityWithoutUnit,
  );

  addIssueSection(
    report,
    "15. Поставщик 3: V есть, но U отсутствует",
    supplier3UnitWithoutQuantity,
  );

  // ----------------------------------------------------------
  // ИТОГ
  // ----------------------------------------------------------

  report.push("");
  report.push("================================================");

  report.push("16. ИТОГ");

  report.push("================================================");

  report.push("");

  report.push(`L URL без F: ${mainUrlWithoutQuantity.length}`);

  report.push(`F без L: ${mainQuantityWithoutUrl.length}`);

  report.push(`F без G: ${mainQuantityWithoutUnit.length}`);

  report.push(`G без F: ${mainUnitWithoutQuantity.length}`);

  report.push(`M без настоящего URL: ${parsingEnabledWithoutUrl.length}`);

  report.push(`I/K без настоящего URL: ${parsedPriceWithoutUrl.length}`);

  report.push(`T без R: ${supplier2UrlWithoutQuantity.length}`);

  report.push(`R без T: ${supplier2QuantityWithoutUrl.length}`);

  report.push(`R без S: ${supplier2QuantityWithoutUnit.length}`);

  report.push(`S без R: ${supplier2UnitWithoutQuantity.length}`);

  report.push(`W без U: ${supplier3UrlWithoutQuantity.length}`);

  report.push(`U без W: ${supplier3QuantityWithoutUrl.length}`);

  report.push(`U без V: ${supplier3QuantityWithoutUnit.length}`);

  report.push(`V без U: ${supplier3UnitWithoutQuantity.length}`);

  report.push("");

  report.push("PostgreSQL НЕ изменялся.");

  report.push("Excel НЕ изменялся.");

  // ----------------------------------------------------------
  // СОХРАНЯЕМ ОТЧЁТ
  // ----------------------------------------------------------

  fs.mkdirSync(reportsDirectory, {
    recursive: true,
  });

  fs.writeFileSync(reportFilePath, report.join("\r\n"), "utf8");

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
// ЗАПУСК
// ------------------------------------------------------------

try {
  validateOffers();
} catch (error) {
  console.error("");
  console.error("Ошибка проверки supplier offers:");
  console.error(error);
  console.error("");

  process.exit(1);
}

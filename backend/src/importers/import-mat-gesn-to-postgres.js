// ============================================================
// PRICE CONTROL V3
// РЕАЛЬНЫЙ ИМПОРТ ЛИСТА "МАТ ГЭСН" В POSTGRESQL
//
// Этот скрипт является первым реальным загрузчиком базы
// материалов Price Control V3.
//
// Что делает скрипт:
//
// 1. Читает Excel-лист "МАТ ГЭСН".
// 2. Проверяет и нормализует исходные данные.
// 3. Создаёт справочник известных поставщиков.
// 4. Создаёт материалы.
// 5. Создаёт предложения поставщиков.
// 6. Сохраняет исторические цены и формулы.
// 7. Создаёт подробный отчёт.
// 8. Выполняет всю запись одной транзакцией.
//
// ВАЖНО:
//
// Если во время записи произойдёт критическая ошибка,
// выполняется ROLLBACK.
//
// То есть база не должна остаться загруженной наполовину.
//
// ДОПОЛНИТЕЛЬНАЯ ЗАЩИТА:
//
// - разрешена только база price_control_v3_dev;
// - требуется явный параметр --execute;
// - основные таблицы перед импортом должны быть пустыми;
// - контрольные количества должны совпасть с dry-run.
// ============================================================

const fs = require("fs");
const path = require("path");

const XLSX = require("xlsx");
const { Client } = require("pg");
const dotenv = require("dotenv");

// ============================================================
// 1. ЗАГРУЖАЕМ .ENV BACKEND
// ============================================================
//
// Скрипт находится:
//
// backend/src/importers/
//
// Поэтому ../../.env приводит нас к:
//
// backend/.env
// ============================================================

dotenv.config({
  path: path.resolve(__dirname, "../../.env"),
});

// ============================================================
// 2. ОСНОВНЫЕ НАСТРОЙКИ
// ============================================================

const TARGET_SHEET_NAME = "МАТ ГЭСН";

// ------------------------------------------------------------
// Импорт разрешён ТОЛЬКО в DEV-базу.
//
// Даже если по ошибке в .env окажется другая база,
// скрипт остановится.
// ------------------------------------------------------------

const ALLOWED_DATABASE = "price_control_v3_dev";

// ------------------------------------------------------------
// Для реальной записи пользователь обязан явно добавить:
//
// --execute
//
// Без него PostgreSQL не изменяется.
// ------------------------------------------------------------

const EXECUTE_FLAG = "--execute";

// ============================================================
// 3. КОНТРОЛЬНЫЕ ЗНАЧЕНИЯ
//
// Эти числа получены после финального dry-run.
//
// Если исходный Excel изменится или логика подготовки данных
// неожиданно даст другой результат, импорт остановится ДО
// записи в PostgreSQL.
// ============================================================

const EXPECTED = {
  materials: 7334,

  priority1: 4226,
  priority2: 345,
  priority3: 86,

  totalOffers: 4657,

  skippedWithoutName: 1,

  parsingRequested: 1720,
  parsingEnabled: 1716,
};

// ============================================================
// 4. ИЗВЕСТНЫЕ ПОСТАВЩИКИ
//
// Эти записи будут созданы в таблице suppliers.
//
// parser_key — внутренний стабильный код V3.
// В дальнейшем к нему можно будет подключить конкретный
// Playwright-парсер.
//
// enabled=true означает, что поставщик используется системой.
// ============================================================

const KNOWN_SUPPLIERS = [
  {
    code: "SATURN",
    name: "Сатурн",
    parserKey: "saturn",
    domains: ["saturn.net"],
  },

  {
    code: "SDVOR",
    name: "Строительный двор",
    parserKey: "sdvor",
    domains: ["sdvor.com"],
  },

  {
    code: "ETM",
    name: "ЭТМ",
    parserKey: "etm",
    domains: ["etm.ru"],
  },

  {
    code: "TINKO",
    name: "Тинко",
    parserKey: "tinko",
    domains: ["tinko.ru"],
  },

  {
    code: "SANTECH",
    name: "Сантехкомплект",
    parserKey: "santech",
    domains: ["santech.ru"],
  },

  {
    code: "SPK",
    name: "СПК",
    parserKey: "spk",
    domains: ["spk.ru"],
  },

  {
    code: "LUNDA",
    name: "Лунда",
    parserKey: "lunda",
    domains: ["lunda.ru"],
  },

  {
    code: "LINOLEUM",
    name: "Линолеум.ру",
    parserKey: "linoleum",
    domains: ["linoleum.ru"],
  },

  {
    code: "AVANGARD",
    name: "Авангард",
    parserKey: "avangard",
    domains: ["avangard-ekaterinburg.ru", "avangardrf.ru"],
  },

  {
    code: "UKAR",
    name: "UKAR",
    parserKey: "ukar",
    domains: ["ukar.su"],
  },

  {
    code: "PALITRA",
    name: "Палитра",
    parserKey: "palitra",
    domains: ["zavod-palitra.ru"],
  },

  {
    code: "KIRELIS",
    name: "Kirelis",
    parserKey: "kirelis",
    domains: ["kirelis.ru"],
  },

  {
    code: "KRIOGEN",
    name: "Криоген",
    parserKey: "kriogen",
    domains: ["kriogen.ru"],
  },

  {
    code: "VKOMPLEKT",
    name: "VKOMPLEKT",
    parserKey: "vkomplekt",
    domains: ["vkomplekt.spb.ru"],
  },

  {
    code: "DN",
    name: "DN",
    parserKey: "dn",
    domains: ["dn.ru"],
  },

  {
    code: "STOKING",
    name: "STOKING",
    parserKey: "stoking",
    domains: ["stoking.ru"],
  },

  {
    code: "METALLPROFIL",
    name: "Металл Профиль",
    parserKey: "metallprofil",
    domains: ["metallprofil.ru"],
  },

  {
    code: "VSEINSTRUMENTI",
    name: "ВсеИнструменты.ру",
    parserKey: "vseinstrumenti",
    domains: ["vseinstrumenti.ru"],
  },

  {
    code: "PETROVICH",
    name: "Петрович",
    parserKey: "petrovich",
    domains: ["petrovich.ru"],
  },

  {
    code: "LEMANA",
    name: "Лемана ПРО",
    parserKey: "lemana",
    domains: ["lemanapro.ru", "leroymerlin.ru"],
  },
];

// ============================================================
// 5. АРГУМЕНТЫ КОМАНДНОЙ СТРОКИ
// ============================================================

const commandArguments = process.argv.slice(2);

const executeRequested = commandArguments.includes(EXECUTE_FLAG);

// ------------------------------------------------------------
// Первый аргумент, который не начинается с "--",
// считаем путём к Excel.
// ------------------------------------------------------------

const excelFilePath = commandArguments.find(
  (argument) => !argument.startsWith("--"),
);

// ============================================================
// 6. ПУТИ ДЛЯ ОТЧЁТА
// ============================================================

const reportsDirectory = path.resolve(__dirname, "../../reports");

function makeTimestamp() {
  const now = new Date();

  const pad = (value) => String(value).padStart(2, "0");

  return (
    `${now.getFullYear()}` +
    `${pad(now.getMonth() + 1)}` +
    `${pad(now.getDate())}-` +
    `${pad(now.getHours())}` +
    `${pad(now.getMinutes())}` +
    `${pad(now.getSeconds())}`
  );
}

const reportFilePath = path.join(
  reportsDirectory,
  `mat-gesn-import-${makeTimestamp()}.txt`,
);

// ============================================================
// 7. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ EXCEL
// ============================================================

function getCell(worksheet, column, rowNumber) {
  return worksheet[`${column}${rowNumber}`];
}

// ------------------------------------------------------------
// Получаем отображаемое значение ячейки.
// ------------------------------------------------------------

function displayCell(cell) {
  if (!cell) {
    return "";
  }

  if (cell.w !== undefined && cell.w !== null && String(cell.w).trim() !== "") {
    return String(cell.w).trim();
  }

  if (cell.v === undefined || cell.v === null) {
    return "";
  }

  if (cell.v instanceof Date) {
    return cell.v.toISOString();
  }

  return String(cell.v).trim();
}

function cellHasValue(cell) {
  return displayCell(cell) !== "";
}

function hasAnyValue(cells) {
  return cells.some(cellHasValue);
}

// ============================================================
// 8. РАБОТА С ПРИМЕЧАНИЯМИ
//
// Несколько технических замечаний объединяются в одно поле
// notes через перевод строки.
// ============================================================

function addNote(notes, text) {
  if (!text) {
    return;
  }

  notes.push(text);
}

function joinNotes(notes) {
  if (!notes.length) {
    return null;
  }

  return notes.join("\n");
}

// ============================================================
// 9. ПРЕДУПРЕЖДЕНИЯ ИМПОРТА
// ============================================================

function addWarning(state, type, rowNumber, resourceCode, message) {
  state.warnings.push({
    type,
    rowNumber,
    resourceCode: resourceCode || "",
    message,
  });

  state.warningCounts[type] = (state.warningCounts[type] || 0) + 1;
}

// ============================================================
// 10. РАЗБОР ЧИСЕЛ
// ============================================================

function parseFlexibleNumber(cell, options = {}) {
  const { repairTrailingSeparator = false } = options;

  if (!cell || !cellHasValue(cell)) {
    return {
      status: "empty",
      value: null,
      raw: "",
      repaired: false,
    };
  }

  // ----------------------------------------------------------
  // Excel-ошибка.
  // ----------------------------------------------------------

  if (cell.t === "e") {
    return {
      status: "invalid",
      value: null,
      raw: displayCell(cell),
      repaired: false,
    };
  }

  // ----------------------------------------------------------
  // Настоящее числовое значение Excel.
  // ----------------------------------------------------------

  if (typeof cell.v === "number" && Number.isFinite(cell.v)) {
    return {
      status: "number",
      value: cell.v,
      raw: displayCell(cell),
      repaired: false,
    };
  }

  // ----------------------------------------------------------
  // TRUE / FALSE не являются числом.
  // ----------------------------------------------------------

  if (typeof cell.v === "boolean") {
    return {
      status: "invalid",
      value: null,
      raw: String(cell.v).toUpperCase(),
      repaired: false,
    };
  }

  const originalText = displayCell(cell);

  let valueText = originalText.replace(/\u00A0/g, " ").trim();

  if (!valueText) {
    return {
      status: "empty",
      value: null,
      raw: "",
      repaired: false,
    };
  }

  const invalidMarkers = new Set([
    "-",
    "—",
    "#VALUE!",
    "#ERROR!",
    "#REF!",
    "FALSE",
    "TRUE",
  ]);

  if (invalidMarkers.has(valueText.toUpperCase())) {
    return {
      status: "invalid",
      value: null,
      raw: originalText,
      repaired: false,
    };
  }

  valueText = valueText.replace(/\s+/g, "");

  let repaired = false;

  // ----------------------------------------------------------
  // Специальная безопасная коррекция количества.
  //
  // В старом Excel найдено:
  //
  // 4,5,
  //
  // Это очевидная лишняя запятая в конце.
  //
  // Такая коррекция разрешается только там,
  // где вызывающий код явно передал:
  //
  // repairTrailingSeparator: true
  // ----------------------------------------------------------

  if (repairTrailingSeparator && /[.,]$/.test(valueText)) {
    const withoutLastCharacter = valueText.slice(0, -1);

    if (/^[+-]?\d+(?:[.,]\d+)?$/.test(withoutLastCharacter)) {
      valueText = withoutLastCharacter;

      repaired = true;
    }
  }

  const lastComma = valueText.lastIndexOf(",");

  const lastDot = valueText.lastIndexOf(".");

  // ----------------------------------------------------------
  // Есть и "," и ".".
  //
  // 1,019.00
  //      → 1019.00
  //
  // 1.019,00
  //      → 1019.00
  // ----------------------------------------------------------

  if (lastComma !== -1 && lastDot !== -1) {
    if (lastComma > lastDot) {
      valueText = valueText.replace(/\./g, "").replace(",", ".");
    } else {
      valueText = valueText.replace(/,/g, "");
    }
  }

  // ----------------------------------------------------------
  // Только запятая.
  // ----------------------------------------------------------
  else if (lastComma !== -1) {
    if (/^\d{1,3}(,\d{3})+$/.test(valueText)) {
      valueText = valueText.replace(/,/g, "");
    } else {
      valueText = valueText.replace(",", ".");
    }
  }

  if (!/^[+-]?\d+(?:\.\d+)?$/.test(valueText)) {
    return {
      status: "invalid",
      value: null,
      raw: originalText,
      repaired,
    };
  }

  const numberValue = Number(valueText);

  if (!Number.isFinite(numberValue)) {
    return {
      status: "invalid",
      value: null,
      raw: originalText,
      repaired,
    };
  }

  return {
    status: "number",
    value: numberValue,
    raw: originalText,
    repaired,
  };
}

// ============================================================
// 11. РАЗБОР ДАТ
// ============================================================

// ============================================================
// 11. РАЗБОР ДАТ
// ============================================================

// ------------------------------------------------------------
// Допустимый диапазон годов для дат обновления цены.
//
// В МАТ ГЭСН даты относятся к истории проверки/парсинга цен.
// Значения вроде 20026 или 20266 являются повреждёнными.
//
// Мы НЕ пытаемся угадывать, что пользователь имел в виду.
// Такие даты возвращаются как invalid и позже сохраняются:
//
// approved_at / last_parsed_at = NULL
//
// а исходное значение попадает в notes и отчёт импорта.
// ------------------------------------------------------------

function isAllowedImportYear(year) {
  return Number.isInteger(year) && year >= 2000 && year <= 2100;
}

// ------------------------------------------------------------
// Создаём корректную UTC-дату.
//
// Дополнительно проверяем:
// - диапазон года;
// - существование месяца;
// - существование дня.
//
// Например 31.02.2026 будет отклонено.
// ------------------------------------------------------------

function makeValidUtcDate(year, month, day) {
  if (!isAllowedImportYear(year)) {
    return null;
  }

  const date = new Date(Date.UTC(year, month - 1, day));

  if (
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() !== month - 1 ||
    date.getUTCDate() !== day
  ) {
    return null;
  }

  return date;
}

// ------------------------------------------------------------
// Преобразуем дату Excel в строку YYYY-MM-DD.
//
// ВАЖНО:
//
// Любая подозрительная дата возвращается как:
//
// status = "invalid"
// value  = null
//
// Импорт из-за неё НЕ останавливается.
// ------------------------------------------------------------

function parseDateCell(cell) {
  if (!cell || !cellHasValue(cell)) {
    return {
      status: "empty",
      value: null,
      raw: "",
    };
  }

  // ----------------------------------------------------------
  // XLSX уже преобразовал значение в JavaScript Date.
  //
  // Именно здесь ранее прошли повреждённые годы:
  //
  // 20026
  // 20266
  //
  // Теперь обязательно проверяем год ДО toISOString().
  // ----------------------------------------------------------

  if (cell.v instanceof Date) {
    if (Number.isNaN(cell.v.getTime())) {
      return {
        status: "invalid",
        value: null,
        raw: displayCell(cell),
      };
    }

    const year = cell.v.getUTCFullYear();

    if (!isAllowedImportYear(year)) {
      return {
        status: "invalid",
        value: null,

        // Если Excel не дал отображаемое значение,
        // сохраняем само исходное Date-представление.
        raw: displayCell(cell) || String(cell.v),
      };
    }

    return {
      status: "date",
      value: cell.v.toISOString().slice(0, 10),

      raw: displayCell(cell),
    };
  }

  // ----------------------------------------------------------
  // На случай числовой Excel serial date.
  // ----------------------------------------------------------

  if (typeof cell.v === "number" && Number.isFinite(cell.v)) {
    const parsed = XLSX.SSF.parse_date_code(cell.v);

    if (parsed && parsed.y && parsed.m && parsed.d) {
      const date = makeValidUtcDate(parsed.y, parsed.m, parsed.d);

      if (date) {
        return {
          status: "date",
          value: date.toISOString().slice(0, 10),

          raw: displayCell(cell),
        };
      }
    }

    return {
      status: "invalid",
      value: null,
      raw: displayCell(cell),
    };
  }

  const valueText = displayCell(cell).trim();

  // ----------------------------------------------------------
  // DD.MM.YYYY
  // ----------------------------------------------------------

  let match = valueText.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);

  if (match) {
    const date = makeValidUtcDate(
      Number(match[3]),
      Number(match[2]),
      Number(match[1]),
    );

    if (date) {
      return {
        status: "date",
        value: date.toISOString().slice(0, 10),

        raw: valueText,
      };
    }

    return {
      status: "invalid",
      value: null,
      raw: valueText,
    };
  }

  // ----------------------------------------------------------
  // M/D/YY или M/D/YYYY
  // ----------------------------------------------------------

  match = valueText.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);

  if (match) {
    let year = Number(match[3]);

    if (year < 100) {
      year += 2000;
    }

    const date = makeValidUtcDate(year, Number(match[1]), Number(match[2]));

    if (date) {
      return {
        status: "date",
        value: date.toISOString().slice(0, 10),

        raw: valueText,
      };
    }

    return {
      status: "invalid",
      value: null,
      raw: valueText,
    };
  }

  // ----------------------------------------------------------
  // YYYY-MM-DD
  // ----------------------------------------------------------

  match = valueText.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);

  if (match) {
    const date = makeValidUtcDate(
      Number(match[1]),
      Number(match[2]),
      Number(match[3]),
    );

    if (date) {
      return {
        status: "date",
        value: date.toISOString().slice(0, 10),

        raw: valueText,
      };
    }

    return {
      status: "invalid",
      value: null,
      raw: valueText,
    };
  }

  // ----------------------------------------------------------
  // Всё, что не прошло проверки, считаем повреждённой датой.
  //
  // Ничего автоматически не исправляем.
  // ----------------------------------------------------------

  return {
    status: "invalid",
    value: null,
    raw: valueText,
  };
}

// ============================================================
// 12. URL И ПОСТАВЩИКИ
// ============================================================

function getUrlInfo(cell) {
  const valueText = displayCell(cell);

  if (!valueText) {
    return {
      type: "empty",
      value: null,
      host: null,
    };
  }

  try {
    const url = new URL(valueText);

    if (url.protocol !== "http:" && url.protocol !== "https:") {
      return {
        type: "reference",
        value: valueText,
        host: null,
      };
    }

    const host = url.hostname.toLowerCase().replace(/^www\./, "");

    return {
      type: "url",
      value: valueText,
      host,
    };
  } catch {
    return {
      type: "reference",
      value: valueText,
      host: null,
    };
  }
}

function hostMatchesDomain(host, domain) {
  return host === domain || host.endsWith(`.${domain}`);
}

function identifySupplier(host) {
  if (!host) {
    return null;
  }

  for (const supplier of KNOWN_SUPPLIERS) {
    for (const domain of supplier.domains) {
      if (hostMatchesDomain(host, domain)) {
        return supplier.code;
      }
    }
  }

  return null;
}

// ============================================================
// 13. ФОРМУЛЫ ПЕРЕСЧЁТА
// ============================================================

function normalizeFormula(formula) {
  return String(formula ?? "")
    .toUpperCase()
    .replace(/\$?([A-Z]+)\$?\d+/g, "$1#")
    .replace(/\s+/g, "");
}

function roundFactor(value) {
  return Number(Number(value).toFixed(8));
}

// ------------------------------------------------------------
// Мы НЕ выполняем формулу Excel.
//
// Для простых линейных вариантов коэффициент получаем через:
//
// approved_price / source_price
//
// Сложные формулы только сохраняются как legacy_formula.
// ------------------------------------------------------------

function classifyConversion({ formula, sourcePrice, approvedPrice }) {
  const formulaText = String(formula ?? "").trim();

  // ----------------------------------------------------------
  // Рабочая цена есть, но формулы нет.
  //
  // Значит считаем её ручной.
  // ----------------------------------------------------------

  if (!formulaText) {
    if (approvedPrice !== null) {
      return {
        mode: "manual",
        factor: 1,
        offset: 0,
      };
    }

    return {
      mode: "factor",
      factor: 1,
      offset: 0,
    };
  }

  const normalized = normalizeFormula(formulaText);

  // ----------------------------------------------------------
  // H + число
  // H - число
  // ----------------------------------------------------------

  const offsetMatch = normalized.match(/^H#([+-])(\d+(?:\.\d+)?)$/);

  if (offsetMatch) {
    const sign = offsetMatch[1] === "+" ? 1 : -1;

    return {
      mode: "factor_offset",
      factor: 1,
      offset: sign * Number(offsetMatch[2]),
    };
  }

  // ----------------------------------------------------------
  // Если H вообще не используется,
  // это уже историческая/сложная формула.
  // ----------------------------------------------------------

  if (!normalized.includes("H#")) {
    return {
      mode: "legacy_formula",
      factor: 1,
      offset: 0,
    };
  }

  const cellReferences = normalized.match(/[A-Z]+#/g) || [];

  const allowedReferences = new Set(["H#", "D#", "F#"]);

  const hasUnsupportedReference = cellReferences.some(
    (reference) => !allowedReferences.has(reference),
  );

  if (hasUnsupportedReference) {
    return {
      mode: "legacy_formula",
      factor: 1,
      offset: 0,
    };
  }

  // ----------------------------------------------------------
  // Явные признаки сложной/повреждённой формулы.
  // ----------------------------------------------------------

  if (
    normalized.includes("!") ||
    normalized.includes(":") ||
    normalized.includes("=") ||
    normalized.includes("**") ||
    normalized.includes("#REF")
  ) {
    return {
      mode: "legacy_formula",
      factor: 1,
      offset: 0,
    };
  }

  const withoutReferences = normalized
    .replace(/H#/g, "")
    .replace(/D#/g, "")
    .replace(/F#/g, "");

  // ----------------------------------------------------------
  // Остались буквы.
  //
  // Значит присутствует неизвестная функция или ссылка.
  // ----------------------------------------------------------

  if (/[A-Z_]/.test(withoutReferences)) {
    return {
      mode: "legacy_formula",
      factor: 1,
      offset: 0,
    };
  }

  // ----------------------------------------------------------
  // Сложение/вычитание кроме H +/- число
  // автоматически не интерпретируем.
  // ----------------------------------------------------------

  if (withoutReferences.includes("+") || withoutReferences.includes("-")) {
    return {
      mode: "legacy_formula",
      factor: 1,
      offset: 0,
    };
  }

  // ----------------------------------------------------------
  // Простой линейный пересчёт.
  //
  // Используем уже рассчитанный Excel результат,
  // а не исполняем старую формулу.
  // ----------------------------------------------------------

  if (sourcePrice !== null && sourcePrice !== 0 && approvedPrice !== null) {
    const calculatedFactor = roundFactor(approvedPrice / sourcePrice);

    // --------------------------------------------------------
    // conversion_factor имеет тип numeric(18,8).
    //
    // Для аномально большого коэффициента безопаснее оставить
    // legacy_formula, чем получить ошибку PostgreSQL.
    // --------------------------------------------------------

    if (
      Number.isFinite(calculatedFactor) &&
      Math.abs(calculatedFactor) < 10000000000
    ) {
      return {
        mode: "factor",
        factor: calculatedFactor,
        offset: 0,
      };
    }
  }

  return {
    mode: "legacy_formula",
    factor: 1,
    offset: 0,
  };
}

// ============================================================
// 14. ПОДГОТОВКА ПЛАНА ИМПОРТА
// ============================================================

function buildImportPlan(resolvedExcelPath) {
  const workbook = XLSX.readFile(resolvedExcelPath, {
    cellDates: true,
    cellFormula: true,
  });

  if (!workbook.SheetNames.includes(TARGET_SHEET_NAME)) {
    throw new Error(`Лист "${TARGET_SHEET_NAME}" не найден.`);
  }

  const worksheet = workbook.Sheets[TARGET_SHEET_NAME];

  const range = XLSX.utils.decode_range(worksheet["!ref"]);

  const state = {
    materials: [],

    warnings: [],
    warningCounts: {},

    stats: {
      materials: 0,

      skippedWithoutName: 0,

      priority1: 0,
      priority2: 0,
      priority3: 0,

      parsingRequested: 0,
      parsingEnabled: 0,

      parsingDisabledNoUrl: 0,
      parsingDisabledUnknownSupplier: 0,

      conversionModes: {
        factor: 0,
        factor_offset: 0,
        manual: 0,
        legacy_formula: 0,
      },
    },
  };

  // ==========================================================
  // ОБХОД СТРОК EXCEL
  // ==========================================================

  for (let rowIndex = 1; rowIndex <= range.e.r; rowIndex += 1) {
    const rowNumber = rowIndex + 1;

    const cells = {};

    for (let columnIndex = 0; columnIndex <= 22; columnIndex += 1) {
      const column = XLSX.utils.encode_col(columnIndex);

      cells[column] = getCell(worksheet, column, rowNumber);
    }

    // --------------------------------------------------------
    // Полностью пустую строку пропускаем.
    // --------------------------------------------------------

    if (!hasAnyValue(Object.values(cells))) {
      continue;
    }

    const resourceCode = displayCell(cells.A);

    const officialName = displayCell(cells.B);

    // --------------------------------------------------------
    // materials.official_name имеет NOT NULL.
    //
    // Придумывать наименование нельзя.
    // --------------------------------------------------------

    if (!officialName) {
      state.stats.skippedWithoutName += 1;

      addWarning(
        state,
        "MISSING_OFFICIAL_NAME",
        rowNumber,
        resourceCode,
        "Материал пропущен: отсутствует official_name.",
      );

      continue;
    }

    const materialNotes = [];

    // ========================================================
    // 15. ОСНОВНЫЕ ДАННЫЕ MATERIAL
    // ========================================================

    const baseQuantity = parseFlexibleNumber(cells.D);

    const baseUnit = displayCell(cells.E);

    let materialBaseQuantity = null;

    if (baseQuantity.status === "number") {
      materialBaseQuantity = baseQuantity.value;
    }

    if (baseQuantity.status === "invalid") {
      addNote(
        materialNotes,
        `[Legacy МАТ ГЭСН] D="${baseQuantity.raw}" не распознано как числовое base_quantity.`,
      );

      addWarning(
        state,
        "INVALID_BASE_QUANTITY",
        rowNumber,
        resourceCode,
        `D="${baseQuantity.raw}"`,
      );
    }

    if (baseQuantity.status === "number" && !baseUnit) {
      addNote(
        materialNotes,
        "[Legacy МАТ ГЭСН] Количество D заполнено, но единица E отсутствует.",
      );

      addWarning(
        state,
        "BASE_QUANTITY_WITHOUT_UNIT",
        rowNumber,
        resourceCode,
        `D="${baseQuantity.raw}", E пусто.`,
      );
    }

    if (baseQuantity.status === "empty" && baseUnit) {
      addNote(
        materialNotes,
        `[Legacy МАТ ГЭСН] E="${baseUnit}", но количество D отсутствует.`,
      );

      addWarning(
        state,
        "BASE_UNIT_WITHOUT_QUANTITY",
        rowNumber,
        resourceCode,
        `D пусто, E="${baseUnit}".`,
      );
    }

    const verified = displayCell(cells.Q) === "1";

    // ========================================================
    // 16. ОПРЕДЕЛЯЕМ ОСНОВНОЙ OFFER
    //
    // M / N / O сами по себе offer НЕ создают.
    // ========================================================

    const mainOfferExists = hasAnyValue([
      cells.C,
      cells.F,
      cells.G,
      cells.H,
      cells.I,
      cells.J,
      cells.K,
      cells.L,
    ]);

    const materialPlan = {
      legacy_row_number: rowNumber,

      resource_code: resourceCode || null,

      official_name: officialName,

      base_quantity: materialBaseQuantity,

      base_unit: baseUnit || null,

      verified,

      notes: null,

      offers: [],
    };

    // ========================================================
    // 17. ОСНОВНОЙ OFFER — PRIORITY 1
    // ========================================================

    if (mainOfferExists) {
      state.stats.priority1 += 1;

      const offerNotes = [];

      // ------------------------------------------------------
      // Количество товара поставщика.
      // ------------------------------------------------------

      const supplierQuantity = parseFlexibleNumber(cells.F);

      if (supplierQuantity.status === "invalid") {
        addNote(
          offerNotes,
          `[Legacy МАТ ГЭСН] F="${supplierQuantity.raw}" не распознано как количество поставщика.`,
        );

        addWarning(
          state,
          "INVALID_SUPPLIER_QUANTITY",
          rowNumber,
          resourceCode,
          `F="${supplierQuantity.raw}"`,
        );
      }

      // ------------------------------------------------------
      // Четыре типа цены.
      // ------------------------------------------------------

      const priceH = parseFlexibleNumber(cells.H);

      const priceI = parseFlexibleNumber(cells.I);

      const priceJ = parseFlexibleNumber(cells.J);

      const priceK = parseFlexibleNumber(cells.K);

      const priceChecks = [
        {
          column: "H",
          result: priceH,
          type: "INVALID_SOURCE_PRICE",
        },

        {
          column: "I",
          result: priceI,
          type: "INVALID_PARSED_SOURCE_PRICE",
        },

        {
          column: "J",
          result: priceJ,
          type: "INVALID_APPROVED_PRICE",
        },

        {
          column: "K",
          result: priceK,
          type: "INVALID_PARSED_PRICE",
        },
      ];

      for (const priceCheck of priceChecks) {
        if (priceCheck.result.status === "invalid") {
          addNote(
            offerNotes,
            `[Legacy МАТ ГЭСН] ${priceCheck.column}="${priceCheck.result.raw}" не распознано как цена.`,
          );

          addWarning(
            state,
            priceCheck.type,
            rowNumber,
            resourceCode,
            `${priceCheck.column}="${priceCheck.result.raw}"`,
          );
        }
      }

      // ------------------------------------------------------
      // Даты.
      // ------------------------------------------------------

      const parsedDateN = parseDateCell(cells.N);

      const approvedDateO = parseDateCell(cells.O);

      if (parsedDateN.status === "invalid") {
        addNote(
          offerNotes,
          `[Legacy МАТ ГЭСН] N="${parsedDateN.raw}" не распознано как дата.`,
        );

        addWarning(
          state,
          "INVALID_PARSED_DATE",
          rowNumber,
          resourceCode,
          `N="${parsedDateN.raw}"`,
        );
      }

      if (approvedDateO.status === "invalid") {
        addNote(
          offerNotes,
          `[Legacy МАТ ГЭСН] O="${approvedDateO.raw}" не распознано как дата.`,
        );

        addWarning(
          state,
          "INVALID_APPROVED_DATE",
          rowNumber,
          resourceCode,
          `O="${approvedDateO.raw}"`,
        );
      }

      // ------------------------------------------------------
      // URL / текстовый источник.
      // ------------------------------------------------------

      const sourceInfo = getUrlInfo(cells.L);

      let supplierCode = null;

      if (sourceInfo.type === "url") {
        supplierCode = identifySupplier(sourceInfo.host);
      }

      // ------------------------------------------------------
      // M — старый признак автоматического парсинга.
      //
      // В V3 parsing_enabled=true оставляем только если:
      //
      // 1. M = "да";
      // 2. есть настоящий HTTP/HTTPS URL;
      // 3. поставщик распознан.
      // ------------------------------------------------------

      const parsingRequested =
        displayCell(cells.M).trim().toLowerCase() === "да";

      if (parsingRequested) {
        state.stats.parsingRequested += 1;
      }

      let parsingEnabled = false;

      if (parsingRequested) {
        if (sourceInfo.type !== "url") {
          state.stats.parsingDisabledNoUrl += 1;

          addNote(
            offerNotes,
            '[Legacy МАТ ГЭСН] M="да", но V3 установил parsing_enabled=false: отсутствует HTTP/HTTPS URL.',
          );

          addWarning(
            state,
            "PARSING_DISABLED_NO_URL",
            rowNumber,
            resourceCode,
            `L="${displayCell(cells.L)}"`,
          );
        } else if (!supplierCode) {
          state.stats.parsingDisabledUnknownSupplier += 1;

          addNote(
            offerNotes,
            `[Legacy МАТ ГЭСН] M="да", но V3 установил parsing_enabled=false: поставщик домена "${sourceInfo.host}" пока не сопоставлен.`,
          );

          addWarning(
            state,
            "PARSING_DISABLED_UNKNOWN_SUPPLIER",
            rowNumber,
            resourceCode,
            `domain="${sourceInfo.host}"`,
          );
        } else {
          parsingEnabled = true;

          state.stats.parsingEnabled += 1;
        }
      }

      // ------------------------------------------------------
      // Справка P.
      //
      // Если основной offer существует,
      // P относится к supplier_offers.notes.
      // ------------------------------------------------------

      const noteP = displayCell(cells.P);

      if (noteP) {
        addNote(offerNotes, `[МАТ ГЭСН P] ${noteP}`);
      }

      // ------------------------------------------------------
      // Сохраняем обе старые формулы.
      // ------------------------------------------------------

      const formulaJ = cells.J?.f || "";

      const formulaK = cells.K?.f || "";

      const conversion = classifyConversion({
        formula: formulaJ,

        sourcePrice: priceH.status === "number" ? priceH.value : null,

        approvedPrice: priceJ.status === "number" ? priceJ.value : null,
      });

      state.stats.conversionModes[conversion.mode] += 1;

      materialPlan.offers.push({
        rowNumber,

        priority: 1,

        supplierCode,

        product_name: displayCell(cells.C) || null,

        url: sourceInfo.type === "url" ? sourceInfo.value : null,

        source_reference:
          sourceInfo.type === "reference" ? sourceInfo.value : null,

        supplier_quantity:
          supplierQuantity.status === "number" ? supplierQuantity.value : null,

        supplier_unit: displayCell(cells.G) || null,

        source_price: priceH.status === "number" ? priceH.value : null,

        last_parsed_source_price:
          priceI.status === "number" ? priceI.value : null,

        approved_price: priceJ.status === "number" ? priceJ.value : null,

        last_parsed_price: priceK.status === "number" ? priceK.value : null,

        conversion_factor: conversion.factor,

        conversion_offset: conversion.offset,

        conversion_mode: conversion.mode,

        conversion_formula: formulaJ || null,

        parsed_conversion_formula: formulaK || null,

        parsing_enabled: parsingEnabled,

        approved_at:
          approvedDateO.status === "date" ? approvedDateO.value : null,

        last_parsed_at:
          parsedDateN.status === "date" ? parsedDateN.value : null,

        status: "active",

        notes: joinNotes(offerNotes),
      });
    }

    // ========================================================
    // 18. НЕТ ОСНОВНОГО OFFER
    // ========================================================
    else {
      const M = displayCell(cells.M);

      const N = displayCell(cells.N);

      const O = displayCell(cells.O);

      const P = displayCell(cells.P);

      // ------------------------------------------------------
      // M/N/O без товара/цены/URL не создают пустой offer.
      //
      // Но исходную информацию сохраняем в materials.notes.
      // ------------------------------------------------------

      if (M || N || O) {
        addNote(
          materialNotes,
          `[Legacy МАТ ГЭСН] Служебные данные без supplier_offer: M="${M}", N="${N}", O="${O}".`,
        );

        addWarning(
          state,
          "METADATA_WITHOUT_MAIN_OFFER",
          rowNumber,
          resourceCode,
          `M="${M}", N="${N}", O="${O}"`,
        );
      }

      // ------------------------------------------------------
      // P без основного offer относится к material.
      // ------------------------------------------------------

      if (P) {
        addNote(materialNotes, `[МАТ ГЭСН P] ${P}`);
      }
    }

    // ========================================================
    // 19. ПОСТАВЩИК №2 — PRIORITY 2
    //
    // Offer создаётся только если заполнена T.
    // ========================================================

    const R = displayCell(cells.R);

    const S = displayCell(cells.S);

    const T = displayCell(cells.T);

    if (T) {
      state.stats.priority2 += 1;

      const offerNotes = [];

      const quantity2 = parseFlexibleNumber(cells.R, {
        repairTrailingSeparator: true,
      });

      if (quantity2.repaired) {
        addNote(
          offerNotes,
          `[Legacy МАТ ГЭСН] R="${quantity2.raw}" автоматически нормализовано до ${quantity2.value}.`,
        );

        addWarning(
          state,
          "REPAIRED_SUPPLIER2_QUANTITY",
          rowNumber,
          resourceCode,
          `R="${quantity2.raw}" → ${quantity2.value}`,
        );
      }

      if (quantity2.status === "invalid") {
        addNote(
          offerNotes,
          `[Legacy МАТ ГЭСН] R="${quantity2.raw}" не распознано как количество.`,
        );

        addWarning(
          state,
          "INVALID_SUPPLIER2_QUANTITY",
          rowNumber,
          resourceCode,
          `R="${quantity2.raw}"`,
        );
      }

      const urlInfo2 = getUrlInfo(cells.T);

      const supplierCode2 =
        urlInfo2.type === "url" ? identifySupplier(urlInfo2.host) : null;

      materialPlan.offers.push({
        rowNumber,

        priority: 2,

        supplierCode: supplierCode2,

        product_name: null,

        url: urlInfo2.type === "url" ? urlInfo2.value : null,

        source_reference: urlInfo2.type === "reference" ? urlInfo2.value : null,

        supplier_quantity:
          quantity2.status === "number" ? quantity2.value : null,

        supplier_unit: S || null,

        source_price: null,
        last_parsed_source_price: null,

        approved_price: null,
        last_parsed_price: null,

        conversion_factor: 1,
        conversion_offset: 0,
        conversion_mode: "factor",

        conversion_formula: null,
        parsed_conversion_formula: null,

        parsing_enabled: false,

        approved_at: null,
        last_parsed_at: null,

        status: "active",

        notes: joinNotes(offerNotes),
      });
    }

    // --------------------------------------------------------
    // R/S заполнены, но T отсутствует.
    //
    // Offer не создаём.
    // Информацию сохраняем в material.notes.
    // --------------------------------------------------------
    else if (R || S) {
      addNote(
        materialNotes,
        `[Legacy МАТ ГЭСН] Данные поставщика №2 без URL: R="${R}", S="${S}", T пусто.`,
      );

      addWarning(
        state,
        "SUPPLIER2_DATA_WITHOUT_URL",
        rowNumber,
        resourceCode,
        `R="${R}", S="${S}"`,
      );
    }

    // ========================================================
    // 20. ПОСТАВЩИК №3 — PRIORITY 3
    // ========================================================

    const U = displayCell(cells.U);

    const V = displayCell(cells.V);

    const W = displayCell(cells.W);

    if (W) {
      state.stats.priority3 += 1;

      const offerNotes = [];

      const quantity3 = parseFlexibleNumber(cells.U, {
        repairTrailingSeparator: true,
      });

      if (quantity3.repaired) {
        addNote(
          offerNotes,
          `[Legacy МАТ ГЭСН] U="${quantity3.raw}" автоматически нормализовано до ${quantity3.value}.`,
        );

        addWarning(
          state,
          "REPAIRED_SUPPLIER3_QUANTITY",
          rowNumber,
          resourceCode,
          `U="${quantity3.raw}" → ${quantity3.value}`,
        );
      }

      if (quantity3.status === "invalid") {
        addNote(
          offerNotes,
          `[Legacy МАТ ГЭСН] U="${quantity3.raw}" не распознано как количество.`,
        );

        addWarning(
          state,
          "INVALID_SUPPLIER3_QUANTITY",
          rowNumber,
          resourceCode,
          `U="${quantity3.raw}"`,
        );
      }

      const urlInfo3 = getUrlInfo(cells.W);

      const supplierCode3 =
        urlInfo3.type === "url" ? identifySupplier(urlInfo3.host) : null;

      materialPlan.offers.push({
        rowNumber,

        priority: 3,

        supplierCode: supplierCode3,

        product_name: null,

        url: urlInfo3.type === "url" ? urlInfo3.value : null,

        source_reference: urlInfo3.type === "reference" ? urlInfo3.value : null,

        supplier_quantity:
          quantity3.status === "number" ? quantity3.value : null,

        supplier_unit: V || null,

        source_price: null,
        last_parsed_source_price: null,

        approved_price: null,
        last_parsed_price: null,

        conversion_factor: 1,
        conversion_offset: 0,
        conversion_mode: "factor",

        conversion_formula: null,
        parsed_conversion_formula: null,

        parsing_enabled: false,

        approved_at: null,
        last_parsed_at: null,

        status: "active",

        notes: joinNotes(offerNotes),
      });
    } else if (U || V) {
      addNote(
        materialNotes,
        `[Legacy МАТ ГЭСН] Данные поставщика №3 без URL: U="${U}", V="${V}", W пусто.`,
      );

      addWarning(
        state,
        "SUPPLIER3_DATA_WITHOUT_URL",
        rowNumber,
        resourceCode,
        `U="${U}", V="${V}"`,
      );
    }

    // ========================================================
    // 21. ЗАВЕРШАЕМ MATERIAL
    // ========================================================

    materialPlan.notes = joinNotes(materialNotes);

    state.materials.push(materialPlan);

    state.stats.materials += 1;
  }

  return state;
}

// ============================================================
// 22. КОНТРОЛЬ ПЛАНА
// ============================================================

function validatePlan(state) {
  const actualTotalOffers =
    state.stats.priority1 + state.stats.priority2 + state.stats.priority3;

  const checks = [
    ["materials", state.stats.materials, EXPECTED.materials],

    ["priority=1", state.stats.priority1, EXPECTED.priority1],

    ["priority=2", state.stats.priority2, EXPECTED.priority2],

    ["priority=3", state.stats.priority3, EXPECTED.priority3],

    ["supplier_offers", actualTotalOffers, EXPECTED.totalOffers],

    [
      "пропущено без названия",
      state.stats.skippedWithoutName,
      EXPECTED.skippedWithoutName,
    ],

    ['M="да"', state.stats.parsingRequested, EXPECTED.parsingRequested],

    [
      "parsing_enabled=true",
      state.stats.parsingEnabled,
      EXPECTED.parsingEnabled,
    ],
  ];

  const mismatches = checks.filter(
    ([, actual, expected]) => actual !== expected,
  );

  if (mismatches.length > 0) {
    const details = mismatches
      .map(
        ([name, actual, expected]) =>
          `${name}: получено ${actual}, ожидалось ${expected}`,
      )
      .join("; ");

    throw new Error(`Контрольные количества не совпали. ${details}`);
  }
}

// ============================================================
// 23. КОНФИГУРАЦИЯ POSTGRESQL
// ============================================================

function getDatabaseConfig() {
  // ----------------------------------------------------------
  // Если когда-нибудь будет использоваться DATABASE_URL,
  // поддерживаем и такой вариант.
  // ----------------------------------------------------------

  if (process.env.DATABASE_URL) {
    return {
      connectionString: process.env.DATABASE_URL,
    };
  }

  return {
    host: process.env.DB_HOST || process.env.PGHOST || "127.0.0.1",

    port: Number(process.env.DB_PORT || process.env.PGPORT || 5432),

    database: process.env.DB_NAME || process.env.PGDATABASE,

    user: process.env.DB_USER || process.env.PGUSER,

    password: process.env.DB_PASSWORD || process.env.PGPASSWORD,
  };
}

// ============================================================
// 24. ПРОВЕРКА DEV-БАЗЫ
// ============================================================

async function verifyDatabaseIsSafe(client) {
  const databaseResult = await client.query(`
      SELECT
        current_database() AS database_name,
        current_user AS user_name
    `);

  const databaseName = databaseResult.rows[0].database_name;

  if (databaseName !== ALLOWED_DATABASE) {
    throw new Error(
      `Импорт запрещён. Подключена база "${databaseName}", а разрешена только "${ALLOWED_DATABASE}".`,
    );
  }

  // ----------------------------------------------------------
  // Блокируем рабочие таблицы на время транзакции.
  //
  // Это исключает случайную параллельную запись из интерфейса.
  // ----------------------------------------------------------

  await client.query(`
    LOCK TABLE
      public.materials,
      public.suppliers,
      public.supplier_offers,
      public.gesn_rows
    IN ACCESS EXCLUSIVE MODE
  `);

  // ----------------------------------------------------------
  // Реальный первый импорт разрешаем только в пустые таблицы.
  //
  // pgmigrations специально НЕ проверяется:
  // история миграций должна оставаться.
  // ----------------------------------------------------------

  const countResult = await client.query(`
      SELECT
        (
          SELECT COUNT(*)::integer
          FROM public.materials
        ) AS materials,

        (
          SELECT COUNT(*)::integer
          FROM public.suppliers
        ) AS suppliers,

        (
          SELECT COUNT(*)::integer
          FROM public.supplier_offers
        ) AS supplier_offers,

        (
          SELECT COUNT(*)::integer
          FROM public.gesn_rows
        ) AS gesn_rows
    `);

  const counts = countResult.rows[0];

  const notEmpty =
    Number(counts.materials) !== 0 ||
    Number(counts.suppliers) !== 0 ||
    Number(counts.supplier_offers) !== 0 ||
    Number(counts.gesn_rows) !== 0;

  if (notEmpty) {
    throw new Error(
      "Импорт остановлен: одна или несколько рабочих таблиц уже содержат данные.",
    );
  }

  return {
    databaseName,

    userName: databaseResult.rows[0].user_name,

    counts,
  };
}

// ============================================================
// 25. СОЗДАЁМ SUPPLIERS
// ============================================================

async function insertSuppliers(client) {
  const supplierIds = new Map();

  for (const supplier of KNOWN_SUPPLIERS) {
    const result = await client.query(
      `
          INSERT INTO public.suppliers (
            code,
            name,
            parser_key,
            enabled
          )
          VALUES (
            $1,
            $2,
            $3,
            TRUE
          )
          RETURNING id
        `,
      [supplier.code, supplier.name, supplier.parserKey],
    );

    supplierIds.set(supplier.code, result.rows[0].id);
  }

  return supplierIds;
}

// ============================================================
// 26. SQL ДЛЯ MATERIAL
// ============================================================

const INSERT_MATERIAL_SQL = `
  INSERT INTO public.materials (
    legacy_row_number,
    resource_code,
    official_name,
    base_quantity,
    base_unit,
    verified,
    notes
  )
  VALUES (
    $1,
    $2,
    $3,
    $4,
    $5,
    $6,
    $7
  )
  RETURNING id
`;

// ============================================================
// 27. SQL ДЛЯ SUPPLIER_OFFER
// ============================================================

const INSERT_OFFER_SQL = `
  INSERT INTO public.supplier_offers (
    material_id,
    supplier_id,
    priority,

    product_name,
    url,

    supplier_quantity,
    supplier_unit,

    conversion_factor,
    parsing_enabled,

    approved_price,
    last_parsed_price,

    approved_at,
    last_parsed_at,

    status,
    notes,

    source_price,
    last_parsed_source_price,

    conversion_offset,
    conversion_mode,

    conversion_formula,
    parsed_conversion_formula,

    source_reference
  )
  VALUES (
    $1,
    $2,
    $3,

    $4,
    $5,

    $6,
    $7,

    $8,
    $9,

    $10,
    $11,

    $12,
    $13,

    $14,
    $15,

    $16,
    $17,

    $18,
    $19,

    $20,
    $21,

    $22
  )
  RETURNING id
`;

// ============================================================
// 28. ЗАПИСЬ MATERIALS И OFFERS
// ============================================================

async function insertMaterialsAndOffers(client, state, supplierIds) {
  let insertedMaterials = 0;

  let insertedOffers = 0;

  for (const material of state.materials) {
    const materialResult = await client.query(INSERT_MATERIAL_SQL, [
      material.legacy_row_number,

      material.resource_code,

      material.official_name,

      material.base_quantity,

      material.base_unit,

      material.verified,

      material.notes,
    ]);

    const materialId = materialResult.rows[0].id;

    insertedMaterials += 1;

    for (const offer of material.offers) {
      let supplierId = null;

      if (offer.supplierCode) {
        supplierId = supplierIds.get(offer.supplierCode) || null;
      }

      await client.query(INSERT_OFFER_SQL, [
        materialId,
        supplierId,
        offer.priority,

        offer.product_name,
        offer.url,

        offer.supplier_quantity,
        offer.supplier_unit,

        offer.conversion_factor,
        offer.parsing_enabled,

        offer.approved_price,
        offer.last_parsed_price,

        offer.approved_at,
        offer.last_parsed_at,

        offer.status,
        offer.notes,

        offer.source_price,
        offer.last_parsed_source_price,

        offer.conversion_offset,
        offer.conversion_mode,

        offer.conversion_formula,
        offer.parsed_conversion_formula,

        offer.source_reference,
      ]);

      insertedOffers += 1;
    }
  }

  return {
    insertedMaterials,
    insertedOffers,
  };
}

// ============================================================
// 29. ПРОВЕРКА КОЛИЧЕСТВ ПОСЛЕ INSERT
//
// Проверка выполняется ДО COMMIT.
//
// Если количество не совпало,
// вызывается ошибка и вся транзакция откатывается.
// ============================================================

async function verifyInsertedCounts(client) {
  const result = await client.query(`
      SELECT
        (
          SELECT COUNT(*)::integer
          FROM public.materials
        ) AS materials,

        (
          SELECT COUNT(*)::integer
          FROM public.suppliers
        ) AS suppliers,

        (
          SELECT COUNT(*)::integer
          FROM public.supplier_offers
        ) AS supplier_offers,

        (
          SELECT COUNT(*)::integer
          FROM public.gesn_rows
        ) AS gesn_rows,

        (
          SELECT COUNT(*)::integer
          FROM public.supplier_offers
          WHERE priority = 1
        ) AS priority1,

        (
          SELECT COUNT(*)::integer
          FROM public.supplier_offers
          WHERE priority = 2
        ) AS priority2,

        (
          SELECT COUNT(*)::integer
          FROM public.supplier_offers
          WHERE priority = 3
        ) AS priority3,

        (
          SELECT COUNT(*)::integer
          FROM public.supplier_offers
          WHERE parsing_enabled = TRUE
        ) AS parsing_enabled
    `);

  const counts = result.rows[0];

  const checks = [
    ["materials", Number(counts.materials), EXPECTED.materials],

    ["suppliers", Number(counts.suppliers), KNOWN_SUPPLIERS.length],

    ["supplier_offers", Number(counts.supplier_offers), EXPECTED.totalOffers],

    ["priority1", Number(counts.priority1), EXPECTED.priority1],

    ["priority2", Number(counts.priority2), EXPECTED.priority2],

    ["priority3", Number(counts.priority3), EXPECTED.priority3],

    [
      "parsing_enabled",
      Number(counts.parsing_enabled),
      EXPECTED.parsingEnabled,
    ],

    ["gesn_rows", Number(counts.gesn_rows), 0],
  ];

  const mismatches = checks.filter(
    ([, actual, expected]) => actual !== expected,
  );

  if (mismatches.length) {
    throw new Error(
      "Контроль после INSERT не пройден: " +
        mismatches
          .map(
            ([name, actual, expected]) =>
              `${name}=${actual}, ожидалось ${expected}`,
          )
          .join("; "),
    );
  }

  return counts;
}

// ============================================================
// 30. ОТЧЁТ
// ============================================================

function writeReport({
  state,
  resolvedExcelPath,
  status,
  databaseInfo = null,
  inserted = null,
  finalCounts = null,
  error = null,
}) {
  fs.mkdirSync(reportsDirectory, {
    recursive: true,
  });

  const lines = [];

  lines.push("================================================");

  lines.push("PRICE CONTROL V3 — ИМПОРТ МАТ ГЭСН");

  lines.push("================================================");

  lines.push("");

  lines.push(`Статус: ${status}`);

  lines.push(`Excel: ${resolvedExcelPath}`);

  lines.push(`Лист: ${TARGET_SHEET_NAME}`);

  if (databaseInfo) {
    lines.push(`База: ${databaseInfo.databaseName}`);

    lines.push(`Пользователь БД: ${databaseInfo.userName}`);
  }

  lines.push("");

  lines.push(`Подготовлено materials: ${state.stats.materials}`);

  lines.push(`Подготовлено priority=1: ${state.stats.priority1}`);

  lines.push(`Подготовлено priority=2: ${state.stats.priority2}`);

  lines.push(`Подготовлено priority=3: ${state.stats.priority3}`);

  lines.push(
    `Подготовлено supplier_offers: ${
      state.stats.priority1 + state.stats.priority2 + state.stats.priority3
    }`,
  );

  lines.push("");

  lines.push(`M="да": ${state.stats.parsingRequested}`);

  lines.push(`parsing_enabled=true: ${state.stats.parsingEnabled}`);

  lines.push(`Отключено без URL: ${state.stats.parsingDisabledNoUrl}`);

  lines.push(
    `Отключено из-за неизвестного supplier: ${state.stats.parsingDisabledUnknownSupplier}`,
  );

  lines.push("");

  lines.push("Режимы пересчёта:");

  lines.push(`factor: ${state.stats.conversionModes.factor}`);

  lines.push(`factor_offset: ${state.stats.conversionModes.factor_offset}`);

  lines.push(`manual: ${state.stats.conversionModes.manual}`);

  lines.push(`legacy_formula: ${state.stats.conversionModes.legacy_formula}`);

  if (inserted) {
    lines.push("");

    lines.push(`Фактически INSERT materials: ${inserted.insertedMaterials}`);

    lines.push(`Фактически INSERT supplier_offers: ${inserted.insertedOffers}`);
  }

  if (finalCounts) {
    lines.push("");

    lines.push("Контроль БД после INSERT:");

    lines.push(`materials: ${finalCounts.materials}`);

    lines.push(`suppliers: ${finalCounts.suppliers}`);

    lines.push(`supplier_offers: ${finalCounts.supplier_offers}`);

    lines.push(`gesn_rows: ${finalCounts.gesn_rows}`);

    lines.push(`priority1: ${finalCounts.priority1}`);

    lines.push(`priority2: ${finalCounts.priority2}`);

    lines.push(`priority3: ${finalCounts.priority3}`);

    lines.push(`parsing_enabled: ${finalCounts.parsing_enabled}`);
  }

  if (error) {
    lines.push("");

    lines.push("ОШИБКА:");

    lines.push(String(error.stack || error.message || error));
  }

  // ==========================================================
  // СТАТИСТИКА ПРЕДУПРЕЖДЕНИЙ
  // ==========================================================

  lines.push("");

  lines.push("================================================");

  lines.push("ПРЕДУПРЕЖДЕНИЯ");

  lines.push("================================================");

  lines.push("");

  lines.push(`Всего предупреждений: ${state.warnings.length}`);

  lines.push("");

  const warningCountEntries = Object.entries(state.warningCounts).sort(
    (a, b) => b[1] - a[1],
  );

  for (const [type, count] of warningCountEntries) {
    lines.push(`${type}: ${count}`);
  }

  // ==========================================================
  // ПОЛНЫЙ СПИСОК ПРЕДУПРЕЖДЕНИЙ
  // ==========================================================

  lines.push("");

  lines.push("================================================");

  lines.push("ПОЛНЫЙ СПИСОК ПРЕДУПРЕЖДЕНИЙ");

  lines.push("================================================");

  lines.push("");

  for (const warning of state.warnings) {
    lines.push(
      `Строка ${warning.rowNumber}; ` +
        `код "${warning.resourceCode}"; ` +
        `${warning.type}; ` +
        `${warning.message}`,
    );
  }

  fs.writeFileSync(reportFilePath, lines.join("\r\n"), "utf8");
}

// ============================================================
// 31. ОСНОВНАЯ ФУНКЦИЯ
// ============================================================

async function main() {
  // ----------------------------------------------------------
  // Проверяем путь Excel.
  // ----------------------------------------------------------

  if (!excelFilePath) {
    throw new Error("Не указан путь к Excel-файлу.");
  }

  const resolvedExcelPath = path.resolve(excelFilePath);

  if (!fs.existsSync(resolvedExcelPath)) {
    throw new Error(`Excel-файл не найден: ${resolvedExcelPath}`);
  }

  console.log("");
  console.log("PRICE CONTROL V3 — ИМПОРТ МАТ ГЭСН");
  console.log("");

  // ==========================================================
  // СНАЧАЛА ГОТОВИМ ВСЕ ДАННЫЕ В ПАМЯТИ
  //
  // PostgreSQL пока НЕ изменяется.
  // ==========================================================

  console.log("1. Читаем и проверяем Excel...");

  const state = buildImportPlan(resolvedExcelPath);

  console.log("2. Проверяем контрольные количества...");

  validatePlan(state);

  const totalOffers =
    state.stats.priority1 + state.stats.priority2 + state.stats.priority3;

  console.log("");
  console.log(`Материалов подготовлено: ${state.stats.materials}`);

  console.log(`Предложений подготовлено: ${totalOffers}`);

  console.log(`Поставщиков-справочников: ${KNOWN_SUPPLIERS.length}`);

  console.log(`parsing_enabled=true: ${state.stats.parsingEnabled}`);

  console.log("");

  // ==========================================================
  // ГЛАВНЫЙ ПРЕДОХРАНИТЕЛЬ
  //
  // Без --execute на этом месте завершаемся.
  // ==========================================================

  if (!executeRequested) {
    console.log("РЕЖИМ БЕЗ ЗАПИСИ.");

    console.log("Параметр --execute не указан.");

    console.log("PostgreSQL НЕ изменялся.");

    console.log("");

    writeReport({
      state,
      resolvedExcelPath,
      status: "ПЛАН ПРОВЕРЕН — БЕЗ ЗАПИСИ",
    });

    console.log(`Отчёт: ${reportFilePath}`);

    console.log("");

    return;
  }

  // ==========================================================
  // НИЖЕ НАЧИНАЕТСЯ РЕАЛЬНАЯ РАБОТА С POSTGRESQL
  // ==========================================================

  const client = new Client(getDatabaseConfig());

  let transactionStarted = false;

  let databaseInfo = null;

  let inserted = null;

  let finalCounts = null;

  try {
    console.log("3. Подключаемся к PostgreSQL...");

    await client.connect();

    // --------------------------------------------------------
    // Открываем одну транзакцию для ВСЕГО импорта.
    // --------------------------------------------------------

    await client.query("BEGIN");

    transactionStarted = true;

    console.log("4. Проверяем DEV-базу...");

    databaseInfo = await verifyDatabaseIsSafe(client);

    console.log(`   База: ${databaseInfo.databaseName}`);

    console.log("5. Создаём справочник suppliers...");

    const supplierIds = await insertSuppliers(client);

    console.log("6. Создаём materials и supplier_offers...");

    inserted = await insertMaterialsAndOffers(client, state, supplierIds);

    console.log("7. Проверяем результаты ДО COMMIT...");

    finalCounts = await verifyInsertedCounts(client);

    // --------------------------------------------------------
    // Только если все проверки прошли,
    // фиксируем транзакцию.
    // --------------------------------------------------------

    await client.query("COMMIT");

    transactionStarted = false;

    writeReport({
      state,
      resolvedExcelPath,
      status: "COMMIT — ИМПОРТ УСПЕШЕН",

      databaseInfo,
      inserted,
      finalCounts,
    });

    console.log("");
    console.log("============================================");

    console.log("ИМПОРТ УСПЕШНО ЗАВЕРШЁН");

    console.log("============================================");

    console.log("");

    console.log(`materials: ${finalCounts.materials}`);

    console.log(`suppliers: ${finalCounts.suppliers}`);

    console.log(`supplier_offers: ${finalCounts.supplier_offers}`);

    console.log(`parsing_enabled: ${finalCounts.parsing_enabled}`);

    console.log("");

    console.log(`Отчёт: ${reportFilePath}`);

    console.log("");
  } catch (error) {
    // --------------------------------------------------------
    // При любой критической ошибке откатываем ВСЮ загрузку.
    // --------------------------------------------------------

    if (transactionStarted) {
      try {
        await client.query("ROLLBACK");

        transactionStarted = false;
      } catch (rollbackError) {
        console.error("Дополнительная ошибка при ROLLBACK:");

        console.error(rollbackError);
      }
    }

    writeReport({
      state,
      resolvedExcelPath,

      status: "ROLLBACK — ИМПОРТ ОТМЕНЁН",

      databaseInfo,
      inserted,
      finalCounts,
      error,
    });

    console.error("");
    console.error("ИМПОРТ НЕ ВЫПОЛНЕН.");

    console.error("Транзакция отменена.");

    console.error("PostgreSQL должен остаться в исходном состоянии.");

    console.error("");

    console.error(error);

    console.error("");

    console.error(`Отчёт: ${reportFilePath}`);

    console.error("");

    process.exitCode = 1;
  } finally {
    try {
      await client.end();
    } catch {
      // Ничего дополнительного не делаем.
    }
  }
}

// ============================================================
// 32. ЗАПУСК
// ============================================================

main().catch((error) => {
  console.error("");
  console.error("Ошибка запуска импортёра:");

  console.error(error);
  console.error("");

  process.exitCode = 1;
});

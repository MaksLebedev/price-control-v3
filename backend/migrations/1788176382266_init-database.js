export const shorthands = undefined;

export const up = (pgm) => {
  // ============================================================
  // 1. MATERIALS
  // Основной каталог материалов — ядро бывшего листа "МАТ ГЭСН"
  // ============================================================

  pgm.createTable("materials", {
    id: {
      type: "bigserial",
      primaryKey: true,
    },

    // Номер строки исходного файла при первоначальном переносе.
    // Нужен для проверки миграции и сопоставления с Google/Excel.
    legacy_row_number: {
      type: "integer",
    },

    // Код ресурса ФССЦ / ФСБЦ и т.п.
    // НЕ является уникальным — дубли допускаются.
    resource_code: {
      type: "text",
    },

    official_name: {
      type: "text",
      notNull: true,
    },

    // Количество, к которому относится единица сборника.
    // Например: 1 м2, 100 м, 1 т.
    base_quantity: {
      type: "numeric(18,6)",
      notNull: true,
      default: 1,
    },

    base_unit: {
      type: "text",
    },

    verified: {
      type: "boolean",
      notNull: true,
      default: false,
    },

    notes: {
      type: "text",
    },

    created_at: {
      type: "timestamptz",
      notNull: true,
      default: pgm.func("current_timestamp"),
    },

    updated_at: {
      type: "timestamptz",
      notNull: true,
      default: pgm.func("current_timestamp"),
    },
  });

  pgm.createIndex("materials", "resource_code");
  pgm.createIndex("materials", "official_name");
  pgm.createIndex("materials", "legacy_row_number");

  // ============================================================
  // 2. SUPPLIERS
  // Справочник поставщиков / магазинов
  // ============================================================

  pgm.createTable("suppliers", {
    id: {
      type: "bigserial",
      primaryKey: true,
    },

    // Внутренний код поставщика.
    // Например PETROVICH, ETM, TINKO.
    code: {
      type: "text",
      notNull: true,
      unique: true,
    },

    // Человеческое наименование.
    name: {
      type: "text",
      notNull: true,
    },

    // Ключ парсера в программном коде.
    // Может отличаться от отображаемого названия.
    parser_key: {
      type: "text",
    },

    enabled: {
      type: "boolean",
      notNull: true,
      default: true,
    },

    created_at: {
      type: "timestamptz",
      notNull: true,
      default: pgm.func("current_timestamp"),
    },

    updated_at: {
      type: "timestamptz",
      notNull: true,
      default: pgm.func("current_timestamp"),
    },
  });

  pgm.createIndex("suppliers", "name");
  pgm.createIndex("suppliers", "parser_key");

  // ============================================================
  // 3. SUPPLIER_OFFERS
  //
  // Связь:
  // материал -> поставщик -> карточка товара -> URL -> цена
  //
  // Именно отсюда в будущем парсер будет получать URL.
  // ============================================================

  pgm.createTable("supplier_offers", {
    id: {
      type: "bigserial",
      primaryKey: true,
    },

    material_id: {
      type: "bigint",
      notNull: true,
      references: "materials",
      onDelete: "CASCADE",
    },

    supplier_id: {
      type: "bigint",
      notNull: true,
      references: "suppliers",
      onDelete: "RESTRICT",
    },

    // Порядок предпочтения поставщика для данного материала.
    priority: {
      type: "integer",
      notNull: true,
      default: 1,
    },

    // Наименование товара именно на сайте поставщика.
    product_name: {
      type: "text",
    },

    // Критически важное поле.
    // URL карточки, которую должен открывать парсер.
    url: {
      type: "text",
    },

    // Количество/упаковка у поставщика.
    supplier_quantity: {
      type: "numeric(18,6)",
    },

    supplier_unit: {
      type: "text",
    },

    // Коэффициент приведения цены магазина
    // к единице измерения МАТ ГЭСН.
    conversion_factor: {
      type: "numeric(18,8)",
      notNull: true,
      default: 1,
    },

    // Нужно ли данный URL обрабатывать парсером.
    parsing_enabled: {
      type: "boolean",
      notNull: true,
      default: false,
    },

    // Текущая рабочая/подтверждённая цена.
    approved_price: {
      type: "numeric(18,4)",
    },

    // Последняя цена, полученная автоматически парсером.
    last_parsed_price: {
      type: "numeric(18,4)",
    },

    // Когда была подтверждена рабочая цена.
    approved_at: {
      type: "timestamptz",
    },

    // Когда парсер последний раз получил цену.
    last_parsed_at: {
      type: "timestamptz",
    },

    status: {
      type: "text",
      notNull: true,
      default: "active",
    },

    notes: {
      type: "text",
    },

    created_at: {
      type: "timestamptz",
      notNull: true,
      default: pgm.func("current_timestamp"),
    },

    updated_at: {
      type: "timestamptz",
      notNull: true,
      default: pgm.func("current_timestamp"),
    },
  });

  pgm.createIndex("supplier_offers", "material_id");
  pgm.createIndex("supplier_offers", "supplier_id");
  pgm.createIndex("supplier_offers", ["material_id", "priority"]);
  pgm.createIndex("supplier_offers", "parsing_enabled");

  // ============================================================
  // 4. GESN_ROWS
  //
  // Одна физическая таблица для существующего листа "ГЭСН".
  //
  // Типы строк:
  // RATE      - сама расценка
  // WORK_STEP - состав работ
  // MATERIAL  - материал/ресурс расценки
  // ============================================================

  pgm.createTable("gesn_rows", {
    id: {
      type: "bigserial",
      primaryKey: true,
    },

    // Исходный номер строки Excel/Google Sheets.
    // Нужен для первоначальной миграции и проверки.
    legacy_row_number: {
      type: "integer",
    },

    // Порядок строки.
    // Позволяет восстановить исходную последовательность ГЭСН.
    sort_order: {
      type: "integer",
      notNull: true,
    },

    row_type: {
      type: "text",
      notNull: true,
    },

    // Для WORK_STEP и MATERIAL указывает,
    // к какой строке RATE относится запись.
    parent_rate_id: {
      type: "bigint",
      references: "gesn_rows",
      onDelete: "CASCADE",
    },

    // Код ГЭСН.
    // В основном используется для строки RATE.
    rate_code: {
      type: "text",
    },

    name: {
      type: "text",
    },

    // Исходное значение количества/расхода как текст.
    //
    // Примеры:
    // "0,04"
    // "1,25"
    // "П"
    quantity_raw: {
      type: "text",
    },

    // Числовое значение, если оно может быть преобразовано.
    quantity_value: {
      type: "numeric(18,8)",
    },

    unit: {
      type: "text",
    },

    // Цена работы для строки RATE.
    work_price: {
      type: "numeric(18,4)",
    },

    work_verified: {
      type: "boolean",
      notNull: true,
      default: false,
    },

    // Связь материала расценки с каталогом МАТ ГЭСН.
    material_id: {
      type: "bigint",
      references: "materials",
      onDelete: "SET NULL",
    },

    // Исходный код материала сохраняем отдельно,
    // даже если material_id пока не найден.
    resource_code: {
      type: "text",
    },

    notes: {
      type: "text",
    },

    created_at: {
      type: "timestamptz",
      notNull: true,
      default: pgm.func("current_timestamp"),
    },

    updated_at: {
      type: "timestamptz",
      notNull: true,
      default: pgm.func("current_timestamp"),
    },
  });

  pgm.addConstraint(
    "gesn_rows",
    "gesn_rows_row_type_check",
    "CHECK (row_type IN ('RATE', 'WORK_STEP', 'MATERIAL'))",
  );

  pgm.createIndex("gesn_rows", "legacy_row_number");
  pgm.createIndex("gesn_rows", "sort_order");
  pgm.createIndex("gesn_rows", "row_type");
  pgm.createIndex("gesn_rows", "parent_rate_id");
  pgm.createIndex("gesn_rows", "rate_code");
  pgm.createIndex("gesn_rows", "resource_code");
  pgm.createIndex("gesn_rows", "material_id");
};

export const down = (pgm) => {
  // Удаляем в обратном порядке зависимостей.

  pgm.dropTable("gesn_rows");
  pgm.dropTable("supplier_offers");
  pgm.dropTable("suppliers");
  pgm.dropTable("materials");
};

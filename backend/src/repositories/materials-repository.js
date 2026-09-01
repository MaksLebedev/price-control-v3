const { pool } = require("../db");

async function getMaterials({ limit = 100, offset = 0 } = {}) {
  const result = await pool.query(
    `
      SELECT
        id,
        legacy_row_number,
        resource_code,
        official_name,
        base_quantity,
        base_unit,
        verified,
        notes,
        created_at,
        updated_at
      FROM materials
      ORDER BY id
      LIMIT $1
      OFFSET $2
    `,
    [limit, offset],
  );

  return result.rows;
}

async function createMaterial({
  legacyRowNumber = null,
  resourceCode = null,
  officialName,
  baseQuantity = 1,
  baseUnit = null,
  verified = false,
  notes = null,
}) {
  const result = await pool.query(
    `
      INSERT INTO materials (
        legacy_row_number,
        resource_code,
        official_name,
        base_quantity,
        base_unit,
        verified,
        notes
      )
      VALUES ($1, $2, $3, $4, $5, $6, $7)
      RETURNING
        id,
        legacy_row_number,
        resource_code,
        official_name,
        base_quantity,
        base_unit,
        verified,
        notes,
        created_at,
        updated_at
    `,
    [
      legacyRowNumber,
      resourceCode,
      officialName,
      baseQuantity,
      baseUnit,
      verified,
      notes,
    ],
  );

  return result.rows[0];
}

async function updateMaterial(id, changes) {
  const fieldMap = {
    legacyRowNumber: "legacy_row_number",
    resourceCode: "resource_code",
    officialName: "official_name",
    baseQuantity: "base_quantity",
    baseUnit: "base_unit",
    verified: "verified",
    notes: "notes",
  };

  const entries = Object.entries(changes).filter(
    ([key, value]) =>
      Object.prototype.hasOwnProperty.call(fieldMap, key) &&
      value !== undefined,
  );

  if (entries.length === 0) {
    return null;
  }

  const setParts = [];
  const values = [];

  entries.forEach(([key, value], index) => {
    setParts.push(`${fieldMap[key]} = $${index + 1}`);
    values.push(value);
  });

  setParts.push("updated_at = current_timestamp");

  values.push(id);

  const result = await pool.query(
    `
      UPDATE materials
      SET ${setParts.join(", ")}
      WHERE id = $${values.length}
      RETURNING
        id,
        legacy_row_number,
        resource_code,
        official_name,
        base_quantity,
        base_unit,
        verified,
        notes,
        created_at,
        updated_at
    `,
    values,
  );

  return result.rows[0] || null;
}

// ------------------------------------------------------------
// УДАЛЕНИЕ МАТЕРИАЛА ИЗ POSTGRESQL
//
// Функция получает внутренний ID материала
// и удаляет соответствующую запись из таблицы materials.
//
// RETURNING нужен для того, чтобы PostgreSQL вернул
// данные удалённого материала.
//
// Это позволяет backend понять:
// - материал действительно существовал и был удалён;
// - либо записи с таким ID вообще не было.
// ------------------------------------------------------------
async function deleteMaterial(id) {
  const result = await pool.query(
    `
      DELETE FROM materials
      WHERE id = $1
      RETURNING
        id,
        resource_code,
        official_name
    `,
    [id],
  );

  // Если запись была найдена и удалена,
  // возвращаем её данные.
  //
  // Если такого ID не существовало,
  // PostgreSQL вернёт пустой массив и мы возвращаем null.
  return result.rows[0] || null;
}

// ------------------------------------------------------------
// ЭКСПОРТ ФУНКЦИЙ РЕПОЗИТОРИЯ
//
// Эти функции используются в server.js.
// Именно через них Express работает с таблицей materials.
// ------------------------------------------------------------
module.exports = {
  getMaterials,
  createMaterial,
  updateMaterial,
  deleteMaterial,
};

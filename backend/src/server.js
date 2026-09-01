const express = require("express");
const cors = require("cors");

const { testDatabaseConnection } = require("./db");

// ------------------------------------------------------------
// РАБОТА С МАТЕРИАЛАМИ В POSTGRESQL
//
// Server.js принимает HTTP-запросы,
// а непосредственную работу с таблицей materials
// выполняет отдельный repository.
// ------------------------------------------------------------
const {
  getMaterials,
  createMaterial,
  updateMaterial,
  deleteMaterial,
} = require("./repositories/materials-repository");

const app = express();

const PORT = 3001;
const HOST = "0.0.0.0";

app.use(cors());
app.use(express.json());

app.get("/api/status", (req, res) => {
  res.json({
    ok: true,
    service: "Price Control V3 API",
    message: "Price Control V3 API работает",
  });
});

app.get("/api/db/status", async (req, res) => {
  try {
    const info = await testDatabaseConnection();

    res.json({
      ok: true,
      message: "PostgreSQL подключен",
      database: info.database,
      user: info.user,
      port: info.port,
    });
  } catch (error) {
    console.error("Ошибка подключения к PostgreSQL:", error);

    res.status(500).json({
      ok: false,
      message: "Ошибка подключения к PostgreSQL",
    });
  }
});

app.get("/api/materials", async (req, res) => {
  try {
    const materials = await getMaterials();

    res.json({
      ok: true,
      count: materials.length,
      materials,
    });
  } catch (error) {
    console.error("Ошибка получения материалов:", error);

    res.status(500).json({
      ok: false,
      message: "Ошибка получения материалов",
    });
  }
});

app.post("/api/materials", async (req, res) => {
  try {
    const {
      legacyRowNumber,
      resourceCode,
      officialName,
      baseQuantity,
      baseUnit,
      verified,
      notes,
    } = req.body;

    if (!officialName || !officialName.trim()) {
      return res.status(400).json({
        ok: false,
        message: "Не указано наименование материала",
      });
    }

    const material = await createMaterial({
      legacyRowNumber,
      resourceCode,
      officialName: officialName.trim(),
      baseQuantity,
      baseUnit,
      verified,
      notes,
    });

    res.status(201).json({
      ok: true,
      material,
    });
  } catch (error) {
    console.error("Ошибка создания материала:", error);

    res.status(500).json({
      ok: false,
      message: "Ошибка создания материала",
    });
  }
});

app.patch("/api/materials/:id", async (req, res) => {
  try {
    const id = Number(req.params.id);

    if (!Number.isInteger(id) || id <= 0) {
      return res.status(400).json({
        ok: false,
        message: "Некорректный ID материала",
      });
    }

    if (
      req.body.officialName !== undefined &&
      (!req.body.officialName || !req.body.officialName.trim())
    ) {
      return res.status(400).json({
        ok: false,
        message: "Наименование материала не может быть пустым",
      });
    }

    const changes = {
      ...req.body,
    };

    if (typeof changes.officialName === "string") {
      changes.officialName = changes.officialName.trim();
    }

    const material = await updateMaterial(id, changes);

    if (!material) {
      return res.status(404).json({
        ok: false,
        message: "Материал не найден или отсутствуют изменения",
      });
    }

    res.json({
      ok: true,
      material,
    });
  } catch (error) {
    console.error("Ошибка изменения материала:", error);

    res.status(500).json({
      ok: false,
      message: "Ошибка изменения материала",
    });
  }
});

// ------------------------------------------------------------
// DELETE /api/materials/:id
//
// Этот маршрут удаляет материал по его внутреннему ID.
//
// Например:
//
// DELETE /api/materials/15
//
// означает:
// удалить из PostgreSQL материал с ID = 15.
// ------------------------------------------------------------
app.delete("/api/materials/:id", async (req, res) => {
  try {
    // ID приходит из адреса запроса как текст.
    // Преобразуем его в число.
    const id = Number(req.params.id);

    // Проверяем, что пользователь передал нормальный ID.
    //
    // Например, значения "abc", "-5" или "0"
    // не должны попадать в запрос PostgreSQL.
    if (!Number.isInteger(id) || id <= 0) {
      return res.status(400).json({
        ok: false,
        message: "Некорректный ID материала",
      });
    }

    // Просим repository удалить материал из PostgreSQL.
    const material = await deleteMaterial(id);

    // Если PostgreSQL не вернул запись,
    // значит материала с таким ID не существовало.
    if (!material) {
      return res.status(404).json({
        ok: false,
        message: "Материал не найден",
      });
    }

    // Сообщаем React, что удаление прошло успешно.
    res.json({
      ok: true,
      message: "Материал удалён",
      material,
    });
  } catch (error) {
    // Если произошла неожиданная ошибка,
    // записываем техническую информацию в консоль backend.
    console.error("Ошибка удаления материала:", error);

    res.status(500).json({
      ok: false,
      message: "Ошибка удаления материала",
    });
  }
});

app.listen(PORT, HOST, () => {
  console.log(`Price Control V3 API запущен: http://localhost:${PORT}`);
});

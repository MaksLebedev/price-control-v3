const express = require("express");
const cors = require("cors");

const { testDatabaseConnection } = require("./db");

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

app.listen(PORT, HOST, () => {
  console.log(`Price Control V3 API запущен: http://localhost:${PORT}`);
});

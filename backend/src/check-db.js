const { pool, testDatabaseConnection } = require("./db");

async function main() {
  try {
    const info = await testDatabaseConnection();

    console.log("Подключение к PostgreSQL успешно.");
    console.log("База:", info.database);
    console.log("Пользователь:", info.user);
    console.log("Порт:", info.port);
  } catch (error) {
    console.error("Ошибка подключения к PostgreSQL:");
    console.error(error.message);

    process.exitCode = 1;
  } finally {
    await pool.end();
  }
}

main();
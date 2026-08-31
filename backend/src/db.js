const { Pool } = require("pg");
require("dotenv").config();

const pool = new Pool({
  host: process.env.DB_HOST,
  port: Number(process.env.DB_PORT),
  database: process.env.DB_NAME,
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
});

async function testDatabaseConnection() {
  const result = await pool.query(`
    SELECT
      current_database() AS database,
      current_user AS user,
      inet_server_port() AS port
  `);

  return result.rows[0];
}

module.exports = {
  pool,
  testDatabaseConnection,
};

import mysql from "mysql2/promise";
import { env } from "./config.js";

export const pool = mysql.createPool({
  uri: env.DATABASE_URL,
  connectionLimit: 10
});

export async function checkDatabaseConnection() {
  const connection = await pool.getConnection();

  try {
    await connection.ping();
  } finally {
    connection.release();
  }
}


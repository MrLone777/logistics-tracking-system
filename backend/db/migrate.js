import 'dotenv/config';
import fs from 'node:fs/promises';
import pg from 'pg';
const { Pool } = pg;
const pool = new Pool({connectionString: process.env.DATABASE_URL, ssl: process.env.DATABASE_URL?.includes('localhost') ? false : {rejectUnauthorized:false}});
try { const sql = await fs.readFile(new URL('./schema.sql', import.meta.url), 'utf8'); await pool.query(sql); console.log('Database schema is ready.'); }
finally { await pool.end(); }

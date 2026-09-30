import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import pg from 'pg';

const { Pool } = pg;
const app = express();
const port = Number(process.env.PORT || 4000);

const pool = new Pool({ connectionString: process.env.DATABASE_URL, ssl: process.env.DATABASE_URL?.includes('localhost') ? false : { rejectUnauthorized: false } });

app.use(cors({ origin: process.env.FRONTEND_ORIGIN || true }));
app.use(express.json({ limit: '2mb' }));

app.get('/api/health', async (_req, res) => {
  try {
    await pool.query('SELECT 1');
    res.json({ ok: true, database: 'connected' });
  } catch (error) {
    res.status(503).json({ ok: false, database: 'disconnected' });
  }
});

app.get('/api/shipments', async (req, res) => {
  try {
    const { category } = req.query;
    const result = category
      ? await pool.query('SELECT * FROM shipments WHERE category = $1 ORDER BY created_at DESC', [category])
      : await pool.query('SELECT * FROM shipments ORDER BY created_at DESC');
    res.json(result.rows);
  } catch (error) {
    res.status(500).json({ message: 'Failed to load shipments.' });
  }
});

app.post('/api/shipments', async (req, res) => {
  const s = req.body || {};
  const category = String(s.category || '').trim();
  const errors = [];

  if (!['Local', 'Domestic', 'International'].includes(category)) errors.push('Invalid category.');
  if (category === 'Local' && String(s.shipment_type || '').toLowerCase() === 'porter' && !String(s.crn_no || '').trim()) errors.push('CRN No. is mandatory for Porter.');
  if ((category === 'Local' && String(s.shipment_type || '').toLowerCase() !== 'porter') || category === 'Domestic') {
    if (!String(s.eway_bill_no || '').trim()) errors.push('E-Way Bill No. is mandatory.');
    if (!String(s.delivery_challan_no || '').trim()) errors.push('Delivery Challan No. is mandatory.');
  }
  if (category === 'International' && !String(s.tracking_id || '').trim()) errors.push('Tracking ID is mandatory.');

  if (errors.length) return res.status(400).json({ message: 'Validation failed.', errors });

  try {
    const result = await pool.query(`
      INSERT INTO shipments (
        category, shipment_type, name, crn_no, eway_bill_no, delivery_challan_no,
        tracking_lr_no, tracking_id, gst, amount, origin, destination, status
      ) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,'Pending')
      RETURNING *
    `, [
      category, s.shipment_type || null, s.name || null, s.crn_no || null,
      s.eway_bill_no || null, s.delivery_challan_no || null, s.tracking_lr_no || null,
      s.tracking_id || null, s.gst || null, s.amount === '' || s.amount == null ? null : Number(s.amount),
      s.origin || null, s.destination || null
    ]);
    res.status(201).json(result.rows[0]);
  } catch (error) {
    res.status(500).json({ message: 'Failed to save shipment.' });
  }
});

app.patch('/api/shipments/:id', async (req, res) => {
  const { status } = req.body || {};
  const allowed = ['Pending', 'In Transit', 'Out for Delivery', 'Delivered', 'Completed'];
  if (!allowed.includes(status)) return res.status(400).json({ message: 'Invalid status.' });
  try {
    const result = await pool.query('UPDATE shipments SET status=$1, updated_at=NOW() WHERE id=$2 RETURNING *', [status, req.params.id]);
    if (!result.rowCount) return res.status(404).json({ message: 'Shipment not found.' });
    res.json(result.rows[0]);
  } catch (error) {
    res.status(500).json({ message: 'Failed to update shipment.' });
  }
});

app.listen(port, () => console.log(`Logistics backend running on port ${port}`));

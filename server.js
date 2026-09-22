const express = require('express');
const { Pool } = require('pg');
const cors = require('cors');
const ExcelJS = require('exceljs');
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: { rejectUnauthorized: false }
});

// ============================================
// SECTORES
// ============================================
app.get('/api/sectores', async (req, res) => {
  const result = await pool.query('SELECT * FROM sectores WHERE activo = TRUE ORDER BY nombre');
  res.json(result.rows);
});

// ============================================
// ANTIBIOTICOS
// ============================================
app.get('/api/antibioticos', async (req, res) => {
  const result = await pool.query('SELECT * FROM antibioticos WHERE activo = TRUE ORDER BY nombre_generico');
  res.json(result.rows);
});

app.post('/api/antibioticos', async (req, res) => {
  const { nombre_generico, presentacion, stock_minimo_alerta } = req.body;
  const result = await pool.query(
    'INSERT INTO antibioticos (nombre_generico, presentacion, stock_minimo_alerta) VALUES ($1, $2, $3) RETURNING *',
    [nombre_generico, presentacion, stock_minimo_alerta || 30]
  );
  res.json(result.rows[0]);
});

// ============================================
// DASHBOARD - STOCK EN TIEMPO REAL CON SEMÁFORO
// ============================================
app.get('/api/stock', async (req, res) => {
  const query = `
    SELECT 
      a.id_antibiotico,
      a.nombre_generico,
      a.presentacion,
      a.stock_minimo_alerta,
      COALESCE(SUM(i.cantidad), 0) - COALESCE((
        SELECT SUM(e.cantidad) FROM egresos e WHERE e.id_antibiotico = a.id_antibiotico
      ), 0) AS stock_actual
    FROM antibioticos a
    LEFT JOIN ingresos i ON i.id_antibiotico = a.id_antibiotico
    WHERE a.activo = TRUE
    GROUP BY a.id_antibiotico, a.nombre_generico, a.presentacion, a.stock_minimo_alerta
    ORDER BY a.nombre_generico
  `;
  const result = await pool.query(query);
  
  const data = result.rows.map(row => {
    let semaforo = 'verde';
    if (row.stock_actual <= row.stock_minimo_alerta) semaforo = 'rojo';
    else if (row.stock_actual <= row.stock_minimo_alerta * 1.5) semaforo = 'amarillo';
    return { ...row, semaforo };
  });
  
  res.json(data);
});

// ============================================
// LOTES DISPONIBLES (FEFO - primero vence, primero sale)
// ============================================
app.get('/api/lotes/:id_antibiotico', async (req, res) => {
  const { id_antibiotico } = req.params;
  const query = `
    SELECT 
      i.lote,
      i.fecha_vencimiento,
      i.cantidad - COALESCE((
        SELECT SUM(e.cantidad) FROM egresos e 
        WHERE e.lote = i.lote AND e.id_antibiotico = i.id_antibiotico
      ), 0) AS stock_disponible
    FROM ingresos i
    WHERE i.id_antibiotico = $1
    ORDER BY i.fecha_vencimiento ASC
  `;
  const result = await pool.query(query, [id_antibiotico]);
  const disponibles = result.rows.filter(r => r.stock_disponible > 0);
  res.json(disponibles);
});

// ============================================
// INGRESOS
// ============================================
app.post('/api/ingresos', async (req, res) => {
  const { id_antibiotico, cantidad, lote, fecha_vencimiento, proveedor, numero_remito } = req.body;
  const result = await pool.query(
    `INSERT INTO ingresos (id_antibiotico, cantidad, lote, fecha_vencimiento, proveedor, numero_remito) 
     VALUES ($1, $2, $3, $4, $5, $6) RETURNING *`,
    [id_antibiotico, cantidad, lote, fecha_vencimiento, proveedor, numero_remito]
  );
  res.json(result.rows[0]);
});

// ============================================
// EGRESOS (con validación Poka-Yoke)
// ============================================
app.post('/api/egresos', async (req, res) => {
  const { id_antibiotico, id_sector, cantidad, lote, solicitante, tipo, forzar } = req.body;

  // Validación Poka-Yoke: stock total disponible
  const stockQuery = await pool.query(`
    SELECT COALESCE(SUM(cantidad), 0) - COALESCE((
      SELECT SUM(cantidad) FROM egresos WHERE id_antibiotico = $1
    ), 0) AS disponible
    FROM ingresos WHERE id_antibiotico = $1
  `, [id_antibiotico]);
  
  const disponible = stockQuery.rows[0].disponible;

  if (cantidad > disponible) {
    return res.status(400).json({ error: 'STOCK_INSUFICIENTE', disponible });
  }

  // Validación Poka-Yoke: promedio histórico del sector (+200%)
  if (!forzar) {
    const promedioQuery = await pool.query(`
      SELECT AVG(cantidad) AS promedio FROM egresos 
      WHERE id_antibiotico = $1 AND id_sector = $2
    `, [id_antibiotico, id_sector]);
    
    const promedio = promedioQuery.rows[0].promedio;
    if (promedio && cantidad > promedio * 3) {
      return res.status(200).json({ alerta: 'CANTIDAD_ANOMALA', promedio, requiereConfirmacion: true });
    }
  }

  const result = await pool.query(
    `INSERT INTO egresos (id_antibiotico, id_sector, cantidad, lote, solicitante, tipo) 
     VALUES ($1, $2, $3, $4, $5, $6) RETURNING *`,
    [id_antibiotico, id_sector, cantidad, lote, solicitante, tipo || 'Tratamiento paciente']
  );
  res.json(result.rows[0]);
});

// ============================================
// HISTORIAL DE EGRESOS (auditoría)
// ============================================
app.get('/api/egresos', async (req, res) => {
  const query = `
    SELECT e.*, a.nombre_generico, a.presentacion, s.nombre AS sector
    FROM egresos e
    JOIN antibioticos a ON a.id_antibiotico = e.id_antibiotico
    JOIN sectores s ON s.id_sector = e.id_sector
    ORDER BY e.fecha_egreso DESC
    LIMIT 200
  `;
  const result = await pool.query(query);
  res.json(result.rows);
});

// ============================================
// EXPORTAR A EXCEL
// ============================================
app.get('/api/export/excel', async (req, res) => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Egresos');

  sheet.columns = [
    { header: 'Fecha', key: 'fecha_egreso', width: 20 },
    { header: 'Antibiótico', key: 'nombre_generico', width: 25 },
    { header: 'Presentación', key: 'presentacion', width: 15 },
    { header: 'Sector', key: 'sector', width: 25 },
    { header: 'Cantidad', key: 'cantidad', width: 10 },
    { header: 'Lote', key: 'lote', width: 15 },
    { header: 'Solicitante', key: 'solicitante', width: 15 },
    { header: 'Tipo', key: 'tipo', width: 20 }
  ];

  const query = `
    SELECT e.*, a.nombre_generico, a.presentacion, s.nombre AS sector
    FROM egresos e
    JOIN antibioticos a ON a.id_antibiotico = e.id_antibiotico
    JOIN sectores s ON s.id_sector = e.id_sector
    ORDER BY e.fecha_egreso DESC
  `;
  const result = await pool.query(query);
  result.rows.forEach(row => sheet.addRow(row));

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=egresos_antibioticos.xlsx');
  await workbook.xlsx.write(res);
  res.end();
});

// ============================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor corriendo en puerto ${PORT}`));

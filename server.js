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

const LEAD_TIME_DIAS = 15; // días que tarda en llegar una compra nueva (ajustable)

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
// DASHBOARD - STOCK + CONSUMO PROMEDIO + AUTONOMÍA + RENOVACIÓN
// ============================================
app.get('/api/stock', async (req, res) => {
  const query = `
    SELECT 
      a.id_antibiotico,
      a.nombre_generico,
      a.presentacion,
      a.stock_minimo_alerta,
      COALESCE(ing.total_ingresos, 0) - COALESCE(egr.total_egresos, 0) AS stock_actual,
      COALESCE(egr3m.total_3m, 0) AS consumo_3_meses,
      prox.proximo_vencimiento
    FROM antibioticos a
    LEFT JOIN (
      SELECT id_antibiotico, SUM(cantidad) AS total_ingresos 
      FROM ingresos GROUP BY id_antibiotico
    ) ing ON ing.id_antibiotico = a.id_antibiotico
    LEFT JOIN (
      SELECT id_antibiotico, SUM(cantidad) AS total_egresos 
      FROM egresos GROUP BY id_antibiotico
    ) egr ON egr.id_antibiotico = a.id_antibiotico
    LEFT JOIN (
      SELECT id_antibiotico, SUM(cantidad) AS total_3m 
      FROM egresos 
      WHERE fecha_egreso >= NOW() - INTERVAL '90 days'
      GROUP BY id_antibiotico
    ) egr3m ON egr3m.id_antibiotico = a.id_antibiotico
    LEFT JOIN (
      SELECT id_antibiotico, MIN(fecha_vencimiento) AS proximo_vencimiento
      FROM ingresos
      WHERE fecha_vencimiento >= CURRENT_DATE
      GROUP BY id_antibiotico
    ) prox ON prox.id_antibiotico = a.id_antibiotico
    WHERE a.activo = TRUE
    ORDER BY a.nombre_generico
  `;
  const result = await pool.query(query);
  const hoy = new Date();

  const data = result.rows.map(row => {
    const stock_actual = parseInt(row.stock_actual);
    const consumo_3m = parseInt(row.consumo_3_meses);
    const consumo_promedio_diario = consumo_3m / 90;
    const consumo_promedio_mensual = consumo_3m / 3;

    let semaforo = 'verde';
    if (stock_actual <= row.stock_minimo_alerta) semaforo = 'rojo';
    else if (stock_actual <= row.stock_minimo_alerta * 1.5) semaforo = 'amarillo';

    let dias_autonomia = null;
    let fecha_agotamiento = null;
    let fecha_renovacion_sugerida = null;
    let alerta_renovacion = false;

    if (consumo_promedio_diario > 0) {
      dias_autonomia = Math.floor(stock_actual / consumo_promedio_diario);

      const fAgota = new Date(hoy);
      fAgota.setDate(fAgota.getDate() + dias_autonomia);
      fecha_agotamiento = fAgota.toISOString().split('T')[0];

      const fRenueva = new Date(fAgota);
      fRenueva.setDate(fRenueva.getDate() - LEAD_TIME_DIAS);
      fecha_renovacion_sugerida = fRenueva.toISOString().split('T')[0];

      if (fRenueva <= hoy) alerta_renovacion = true;
    }

    return {
      id_antibiotico: row.id_antibiotico,
      nombre_generico: row.nombre_generico,
      presentacion: row.presentacion,
      stock_minimo_alerta: row.stock_minimo_alerta,
      stock_actual,
      consumo_promedio_mensual: Math.round(consumo_promedio_mensual),
      consumo_promedio_diario: Math.round(consumo_promedio_diario * 10) / 10,
      dias_autonomia,
      fecha_agotamiento,
      fecha_renovacion_sugerida,
      alerta_renovacion,
      proximo_vencimiento: row.proximo_vencimiento,
      semaforo
    };
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

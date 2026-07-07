require('dotenv').config();
const express = require('express');
const { google } = require('googleapis');
const cors = require('cors');
const { Pool } = require('pg');
const multer = require('multer');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

// 🐘 PostgreSQL (Neon) — para el formulario de registro de gestión.
// La cadena vive en la variable de entorno DATABASE_URL (nunca en el código).
const pgPool = process.env.DATABASE_URL
  ? new Pool({ connectionString: process.env.DATABASE_URL, ssl: { rejectUnauthorized: false } })
  : null;

// 📤 Subida de archivos en memoria (para el Excel de ventas). Límite 200 MB.
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 200 * 1024 * 1024 } });

// 🔹 Configuración de autenticaciones por tipo
const googleAuthConfigs = {
  // call: new google.auth.GoogleAuth({
  //   keyFile: 'northern-cubist-454520-q8-1292a8b77330.json',
  //   scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  // }),
  claveUnica: new google.auth.GoogleAuth({
    keyFile: 'ffvv-realzza-campo-07c3f6b5b98f.json',
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  }),
};

// 🔹 Configuración de hojas (puedes agregar más fácilmente)
const sheetsConfigs = {
  call: {
    authKey: 'claveUnica',
    spreadsheetId: '1j3b7k-vD9UzWLqz6JJksm5Vj3dWvtqL4SckMP21II94',
    range: 'Respuestas de formulario 1!A:ZZZ',
  },
  campo: {
    authKey: 'claveUnica',
    spreadsheetId: '10rjPaki_8JIxJAyN96QK1Wr2MPeu0MfgVm6G8yezP-s',
    range: 'Form Responses 1!A:ZZZ',
  },
  postVenta: {
    authKey: 'claveUnica',
    spreadsheetId: '1uJGGD-eLH8But-5rGdPcmgHZQl1tRbdsG5aDrHj-5UU',
    range: 'Form Responses 1!A:ZZZ',
  },
  pvCobranza: {
    authKey: 'claveUnica',
    spreadsheetId: '1-jAzHZamSVRSKur_8nI3RoY7n606syviR1pGO23YavA',
    range: 'Respuestas!A:ZZZ',
  },
  pvControlInterno: {
    authKey: 'claveUnica',
    spreadsheetId: '1g80tGBpZpJxz0C4efKq-DnV4_c2KM09_-SWgB4ZDMqg',
    range: 'Respuestas!A:ZZZ',
  },
  pvCreditos: {
    authKey: 'claveUnica',
    spreadsheetId: '1uwYZ3iulZYottE23bPHrFNahw5QJ2nUh0nsU8hu7TXQ',
    range: 'Respuestas!A:ZZZ',
  },
  pvLogistica: {
    authKey: 'claveUnica',
    spreadsheetId: '1jG0ageh-_985ybeta4DlYSvmpWkujFVxWJKC5Dxp3Zs',
    range: 'Respuestas!A:ZZZ',
  },
  pvOperaciones: {
    authKey: 'claveUnica',
    spreadsheetId: '1v1KJSVaeJtGda7qbZhgqCo2bPeSdGsBtSUvRw4ewiO8',
    range: 'Respuestas!A:ZZZ',
  },
  pvServicioTecnico: {
    authKey: 'claveUnica',
    spreadsheetId: '1XcTPp4BiOqwjeP6m9Y6Ubs0bgsNhPTd3OhC1KxN6yWU',
    range: 'Respuestas!A:ZZZ',
  },
  pvVentas: {
    authKey: 'claveUnica',
    spreadsheetId: '17ZG1N52lg7O7i8a-7D6xbkszaUG-bSQXSUjK3OP-QDs',
    range: 'Respuestas!A:ZZZ',
  },
  kommo: {
    authKey: 'claveUnica',
    spreadsheetId: '18Dcde-XEdMUEMZ3Fekz3gkXpPP7UzPZxpnoMzcZlm-w',
    range: 'Respuestas de formulario 1!A:ZZZ',
  },
  ferre: {
    authKey: 'claveUnica',
    spreadsheetId: '1q8flDOGxiZdhmP3Kpz4m8s74AKb6b8j60hgRlqfNpPo',
    range: 'Respuestas de formulario 1!A:ZZZ',
  },
  sedes: {
    authKey: 'claveUnica',
    spreadsheetId: '1zHH-1n2fxknSOfPBje0U3x_hRXBXBh4M21KRGHGBbbQ',
    range: 'Respuestas de formulario 1!A:ZZZ',
  },
  // CAP de asesores por sede (no es un formulario: es una hoja normal).
  // Se lee de la pestaña "CAP", cuya fila 1 son las cabeceras
  // (VENDEDOR, SEDE, SUPERVISOR, GERENTE DE TIENDA, ZONA, CANAL, ESTADO, TIPO AV).
  capSedes: {
    authKey: 'claveUnica',
    spreadsheetId: '1_mp6v9g6BfWZ4Otbmv2PTcmkicCqq9fhbAihR0CdHgQ',
    range: 'CAP!A:ZZZ',
  },
  usuarios: {
    authKey: 'claveUnica',
    spreadsheetId: '1z7Qx5vvwCkX2TjVbhUIBR8cMCW3IcdAQHXrIalz0_ZI',
    range: 'usuarios!A:F',
  }
};

// 📌 Ruta dinámica: /form/:sheetName
app.get('/data/:sheetName', async (req, res) => {
  const { sheetName } = req.params;
  const config = sheetsConfigs[sheetName];

  if (!config) {
    return res.status(400).json({ error: 'El nombre del formulario no es válido.' });
  }

  try {
    // Autenticación dinámica
    const auth = googleAuthConfigs[config.authKey];
    const client = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: client });

    // Obtener datos
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: config.spreadsheetId,
      range: config.range,
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      return res.status(404).send('No se encontraron datos en Google Sheets.');
    }

    const [rawHeaders, ...data] = rows;

    // Evita encabezados duplicados
    const headers = [];
    const headerCount = {};

    rawHeaders.forEach((header) => {
      if (!headerCount[header]) {
        headerCount[header] = 1;
        headers.push(header);
      } else {
        const newHeader = `${header} (${headerCount[header]})`;
        headerCount[header]++;
        headers.push(newHeader);
      }
    });

    // Transformar filas a JSON
    let jsonData = data.map((row) =>
      headers.reduce((acc, header, i) => ({
        ...acc,
        [header]: row[i] || '',
      }), {})
    );

    // 🔎 Filtro opcional por fecha (evita traer todo el histórico: p.ej. /data/sedes?desde=2026-07-01&hasta=2026-07-31)
    // Se filtra por la columna de fecha del sheet ("Marca temporal" o "Timestamp"),
    // comparando en formato yyyy-mm-dd. Solo aplica si se envía desde/hasta.
    const { desde, hasta } = req.query;
    if (desde || hasta) {
      const colFecha = headers.includes('Marca temporal')
        ? 'Marca temporal'
        : (headers.includes('Timestamp') ? 'Timestamp' : null);
      if (colFecha) {
        const toKey = (marca) => {
          if (!marca) return null;
          const p = String(marca).trim().split(' ')[0].split('/'); // d/M/yyyy
          if (p.length !== 3) return null;
          const [d, mo, y] = p;
          if (!y || !mo || !d) return null;
          return `${y}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
        };
        jsonData = jsonData.filter((row) => {
          const k = toKey(row[colFecha]);
          if (!k) return false;
          if (desde && k < desde) return false;
          if (hasta && k > hasta) return false;
          return true;
        });
      }
    }

    res.json(jsonData);
  } catch (error) {
    console.error(`❌ Error al obtener datos de ${sheetName}:`, error);
    res.status(500).send('Error al obtener datos de Google Sheets');
  }
});

// 🔐 Login: POST /auth/login
app.post('/auth/login', async (req, res) => {
  const { usuario, password } = req.body;

  if (!usuario || !password) {
    return res.status(400).json({ success: false, message: 'Usuario y contraseña requeridos.' });
  }

  try {
    const config = sheetsConfigs['usuarios'];
    const auth = googleAuthConfigs[config.authKey];
    const client = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: client });

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: config.spreadsheetId,
      range: config.range,
    });

    const rows = response.data.values;
    if (!rows || rows.length < 2) {
      return res.status(500).json({ success: false, message: 'No se encontraron usuarios.' });
    }

    const [headers, ...data] = rows;
    const idx = (col) => headers.indexOf(col);

    const user = data.find(
      (row) =>
        row[idx('usuario')] === usuario &&
        row[idx('password')] === password &&
        (row[idx('activo')] || '').toUpperCase() === 'SI'
    );

    if (!user) {
      return res.status(401).json({ success: false, message: 'Credenciales inválidas o usuario inactivo.' });
    }

    res.json({
      success: true,
      nombre: user[idx('nombre')] || '',
      rol: user[idx('rol')] || '',
      sede: user[idx('sede')] || '',
    });
  } catch (error) {
    console.error('❌ Error en /auth/login:', error);
    res.status(500).json({ success: false, message: 'Error al autenticar.' });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// 📝 Registro de gestión de ventas (formulario de vendedores) → PostgreSQL
// POST /gestion  (body con los campos del formulario; los condicionales pueden ir null)
// ─────────────────────────────────────────────────────────────────────────────
app.post('/gestion', async (req, res) => {
  if (!pgPool) {
    return res.status(500).json({ success: false, message: 'Base de datos no configurada (falta DATABASE_URL).' });
  }

  const b = req.body || {};
  // Validación mínima de los campos obligatorios base.
  const requeridos = ['dni_cliente', 'sede', 'asesor', 'tipo_gestion', 'resultado'];
  for (const campo of requeridos) {
    if (!b[campo] || b[campo].toString().trim() === '') {
      return res.status(400).json({ success: false, message: `Falta el campo obligatorio: ${campo}.` });
    }
  }

  const norm = (v) => (v === undefined || v === null || v === '' ? null : v);
  const valorVenta = (b.valor_venta === '' || b.valor_venta === undefined || b.valor_venta === null)
    ? null
    : Number(String(b.valor_venta).replace(/[^0-9.]/g, '')) || null;

  try {
    const q = `INSERT INTO gestion
      (registrado_por, dni_cliente, sede, asesor, tipo_gestion, resultado,
       motivo_contacto, motivo_no_contacto, fecha_compromiso, valor_venta,
       producto_interes, detalle_contacto, celular_actualizado)
      VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13) RETURNING *`;
    const vals = [
      norm(b.registrado_por), b.dni_cliente, b.sede, b.asesor, b.tipo_gestion, b.resultado,
      norm(b.motivo_contacto), norm(b.motivo_no_contacto), norm(b.fecha_compromiso), valorVenta,
      norm(b.producto_interes), norm(b.detalle_contacto), norm(b.celular_actualizado),
    ];
    const { rows } = await pgPool.query(q, vals);
    res.json({ success: true, gestion: rows[0] });
  } catch (error) {
    console.error('❌ Error en POST /gestion:', error);
    res.status(500).json({ success: false, message: 'No se pudo guardar la gestión.' });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// 🗺️ Optimización de rutas: POST /maps/optimizar
//
// El frontend envía la lista completa de puntos. Aquí (y SOLO aquí) usamos la
// API Key de Google Maps —vive en la variable de entorno GOOGLE_MAPS_API_KEY,
// nunca en el cliente— para pedir a la ROUTES API (computeRoutes) el orden
// óptimo de los waypoints intermedios (optimizeWaypointOrder: true →
// routes[0].optimizedIntermediateWaypointIndex).
//
// Nota: usamos la Routes API (routes.googleapis.com) en lugar de la Directions
// API "legacy", que Google ya no habilita en proyectos nuevos.
//
// Body esperado:
//   { "coordenadas": [ { "lat": -6.77, "lng": -79.84, "id": "A", "nombre": "Sede" }, ... ],
//     "travelmode": "driving" }   // driving | walking | bicycling | transit
//
// Respuesta:
//   { success, waypointOrder, puntosOptimizados, distanciaMetros, duracionSegundos }
// ─────────────────────────────────────────────────────────────────────────────
const GOOGLE_MAPS_API_KEY = process.env.GOOGLE_MAPS_API_KEY;

// Mapea el modo del frontend al enum de la Routes API.
const TRAVEL_MODE_MAP = {
  driving: 'DRIVE',
  walking: 'WALK',
  bicycling: 'BICYCLE',
  transit: 'TRANSIT',
};

app.post('/maps/optimizar', async (req, res) => {
  const { coordenadas, travelmode = 'driving' } = req.body || {};

  if (!Array.isArray(coordenadas) || coordenadas.length < 2) {
    return res.status(400).json({
      success: false,
      message: 'Se requieren al menos 2 coordenadas (origen y destino).',
    });
  }
  if (!GOOGLE_MAPS_API_KEY) {
    return res.status(500).json({
      success: false,
      message: 'Falta configurar GOOGLE_MAPS_API_KEY en el servidor.',
    });
  }

  const esValida = (c) =>
    c && typeof c.lat === 'number' && typeof c.lng === 'number' &&
    c.lat >= -90 && c.lat <= 90 && c.lng >= -180 && c.lng <= 180;

  if (!coordenadas.every(esValida)) {
    return res.status(400).json({ success: false, message: 'Hay coordenadas inválidas.' });
  }

  try {
    const waypoint = (c) => ({ location: { latLng: { latitude: c.lat, longitude: c.lng } } });
    const origin = coordenadas[0];
    const destination = coordenadas[coordenadas.length - 1];
    const intermedios = coordenadas.slice(1, -1);

    const body = {
      origin: waypoint(origin),
      destination: waypoint(destination),
      travelMode: TRAVEL_MODE_MAP[travelmode] || 'DRIVE',
    };
    if (intermedios.length > 0) {
      body.intermediates = intermedios.map(waypoint);
      // 🔑 Parámetro crítico que pide a Google reordenar los waypoints.
      body.optimizeWaypointOrder = true;
    }

    const apiRes = await fetch(
      'https://routes.googleapis.com/directions/v2:computeRoutes',
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-Goog-Api-Key': GOOGLE_MAPS_API_KEY,
          'X-Goog-FieldMask':
            'routes.optimizedIntermediateWaypointIndex,routes.distanceMeters,routes.duration',
        },
        body: JSON.stringify(body),
      },
    );
    const data = await apiRes.json();

    if (!apiRes.ok || !data.routes?.length) {
      return res.status(502).json({
        success: false,
        message: `Google Routes: ${data?.error?.status || apiRes.status}${data?.error?.message ? ' — ' + data.error.message : ''}`,
      });
    }

    const route = data.routes[0];
    // Orden óptimo de los intermedios (índices 0-based del tramo intermedio).
    const waypointOrder = route.optimizedIntermediateWaypointIndex || [];

    // Reconstruir la lista completa ya ordenada: origen + intermedios óptimos + destino.
    const intermediosOrdenados = waypointOrder.map((i) => intermedios[i]);
    const puntosOptimizados = [origin, ...intermediosOrdenados, destination];

    // Totales (duration viene como string tipo "1234s").
    const distanciaMetros = route.distanceMeters || 0;
    const duracionSegundos = parseInt(String(route.duration || '0').replace('s', ''), 10) || 0;

    res.json({
      success: true,
      waypointOrder,
      puntosOptimizados,
      distanciaMetros,
      duracionSegundos,
    });
  } catch (error) {
    console.error('❌ Error en /maps/optimizar:', error);
    res.status(500).json({ success: false, message: 'Error al optimizar la ruta.' });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// 🛒 VENTAS (afectaciones) → PostgreSQL (Neon)
//
// Fuente: Excel exportado de la tabla dbo.excel_GenerarVentas_Afectaciones_Comportamiento
// (acumulado; las ventas no se eliminan, solo cambian de estado). Por eso la
// carga es un UPSERT por CodigoCV: si existe → UPDATE, si no → INSERT. Idempotente.
//
// Endpoints:
//   POST /ventas/import   (multipart, campo "archivo")  → carga/actualiza el Excel
//   GET  /ventas/estado                                 → total + última carga (timestamp)
//   GET  /ventas?anio=&mes=&sede=                        → filas para los consumidores
// ─────────────────────────────────────────────────────────────────────────────

// Columnas de la tabla (orden usado en el INSERT). codigo_cv es la PK (upsert).
const VENTAS_COLS = [
  'codigo_cv', 'dia_cv', 'mes_cv', 'anio_cv', 'cliente_venta', 'sede',
  'monto_consolidado', 'cuota_inicial', 'productos', 'cuotas', 'doc_identidad',
  'estado_venta', 'entidad', 'vendedor', 'tipo_credito', 'estado_tipo_producto',
  'dia_af', 'mes_af', 'anio_af',
];

let ventasSchemaLista = false;
async function ensureVentasSchema() {
  if (!pgPool || ventasSchemaLista) return;
  await pgPool.query(`
    CREATE TABLE IF NOT EXISTS ventas (
      codigo_cv            BIGINT       PRIMARY KEY,
      dia_cv               SMALLINT,
      mes_cv               SMALLINT,
      anio_cv              SMALLINT,
      cliente_venta        TEXT,
      sede                 TEXT,
      monto_consolidado    NUMERIC(14,2),
      cuota_inicial        NUMERIC(14,2),
      productos            TEXT,
      cuotas               INTEGER,
      doc_identidad        TEXT,
      estado_venta         TEXT,
      entidad              TEXT,
      vendedor             TEXT,
      tipo_credito         TEXT,
      estado_tipo_producto TEXT,
      dia_af               SMALLINT,
      mes_af               SMALLINT,
      anio_af              SMALLINT,
      fecha_cv  DATE GENERATED ALWAYS AS (
                  make_date(NULLIF(anio_cv,0), NULLIF(mes_cv,0), NULLIF(dia_cv,0))) STORED,
      fecha_af  DATE GENERATED ALWAYS AS (
                  make_date(NULLIF(anio_af,0), NULLIF(mes_af,0), NULLIF(dia_af,0))) STORED,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
      created_at TIMESTAMPTZ NOT NULL DEFAULT now()
    );
    CREATE INDEX IF NOT EXISTS ix_ventas_anio_mes ON ventas (anio_cv, mes_cv);
    CREATE INDEX IF NOT EXISTS ix_ventas_sede     ON ventas (sede);
    CREATE INDEX IF NOT EXISTS ix_ventas_fecha_cv ON ventas (fecha_cv);
    CREATE TABLE IF NOT EXISTS ventas_cargas (
      id           BIGSERIAL PRIMARY KEY,
      cargado_por  TEXT,
      archivo      TEXT,
      filas        INTEGER,
      insertados   INTEGER,
      actualizados INTEGER,
      creado_en    TIMESTAMPTZ NOT NULL DEFAULT now()
    );
    -- Migración: la columna se llamaba tipo_venta; la fuente real es TipoCredito.
    DO $$ BEGIN
      IF EXISTS (SELECT 1 FROM information_schema.columns
                 WHERE table_name='ventas' AND column_name='tipo_venta')
         AND NOT EXISTS (SELECT 1 FROM information_schema.columns
                 WHERE table_name='ventas' AND column_name='tipo_credito')
      THEN ALTER TABLE ventas RENAME COLUMN tipo_venta TO tipo_credito; END IF;
    END $$;
  `);
  ventasSchemaLista = true;
}

// Toma el primer valor no vacío entre varias cabeceras posibles (tolera ñ/acentos).
function pickCol(row, ...keys) {
  for (const k of keys) {
    const v = row[k];
    if (v !== undefined && v !== null && String(v).trim() !== '') return v;
  }
  return null;
}
function toInt(v) {
  if (v === null || v === undefined || v === '') return null;
  const n = parseInt(String(v).replace(/[^0-9-]/g, ''), 10);
  return Number.isFinite(n) ? n : null;
}
function toNum(v) {
  if (v === null || v === undefined || v === '') return null;
  const n = Number(String(v).replace(/,/g, '').replace(/[^0-9.-]/g, ''));
  return Number.isFinite(n) ? n : null;
}
function toStr(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  return s === '' ? null : s;
}

// Convierte una fila cruda del Excel al arreglo ordenado según VENTAS_COLS.
// Devuelve null si no hay CodigoCV (fila inservible para el upsert).
function mapVentaRow(r) {
  const codigo = toInt(pickCol(r, 'CodigoCV', 'codigo_cv', 'CODIGOCV', 'Codigo CV'));
  if (codigo === null) return null;
  return [
    codigo,
    toInt(pickCol(r, 'DiaCV', 'dia_cv')),
    toInt(pickCol(r, 'MesCV', 'mes_cv')),
    toInt(pickCol(r, 'AñoCV', 'AnioCV', 'AnoCV', 'anio_cv')),
    toStr(pickCol(r, 'ClienteVenta', 'cliente_venta')),
    toStr(pickCol(r, 'Sede', 'sede')),
    toNum(pickCol(r, 'MontoConsolidado', 'monto_consolidado')),
    toNum(pickCol(r, 'CuotaInicial', 'cuota_inicial')),
    toStr(pickCol(r, 'Productos', 'productos')),
    toInt(pickCol(r, 'Cuotas', 'cuotas')),
    toStr(pickCol(r, 'DocIdentidad', 'doc_identidad')),
    toStr(pickCol(r, 'EstadoVenta', 'estado_venta')),
    toStr(pickCol(r, 'Entidad', 'entidad')),
    toStr(pickCol(r, 'Vendedor', 'vendedor')),
    toStr(pickCol(r, 'TipoCredito', 'TipoVenta', 'tipo_credito')),
    toStr(pickCol(r, 'EstadoTipoProducto', 'estado_tipo_producto')),
    toInt(pickCol(r, 'DiaAF', 'dia_af')),
    toInt(pickCol(r, 'MesAF', 'mes_af')),
    toInt(pickCol(r, 'AñoAF', 'AnioAF', 'AnoAF', 'anio_af')),
  ];
}

// UPSERT de un lote de filas. RETURNING (xmax = 0) distingue insertadas de actualizadas.
const VENTAS_SET = VENTAS_COLS.slice(1).map(c => `${c} = EXCLUDED.${c}`).join(', ') + ', updated_at = now()';
async function upsertVentasChunk(client, chunk) {
  const params = [];
  const tuples = chunk.map((row, i) => {
    const base = i * VENTAS_COLS.length;
    params.push(...row);
    return '(' + VENTAS_COLS.map((_, j) => `$${base + j + 1}`).join(',') + ')';
  });
  const sql = `INSERT INTO ventas (${VENTAS_COLS.join(',')}) VALUES ${tuples.join(',')}
    ON CONFLICT (codigo_cv) DO UPDATE SET ${VENTAS_SET}
    RETURNING (xmax = 0) AS inserted`;
  const { rows } = await client.query(sql, params);
  let inserted = 0;
  for (const r of rows) if (r.inserted) inserted++;
  return { inserted, updated: rows.length - inserted };
}

// POST /ventas/import — recibe el Excel (multipart, campo "archivo") y hace el upsert.
app.post('/ventas/import', upload.single('archivo'), async (req, res) => {
  if (!pgPool) {
    return res.status(500).json({ success: false, message: 'Base de datos no configurada (falta DATABASE_URL).' });
  }
  if (!req.file) {
    return res.status(400).json({ success: false, message: 'No se recibió archivo (campo "archivo").' });
  }

  try {
    const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, { defval: '' });

    // Dedupe por codigo_cv (la última ocurrencia gana). Evita el error de Postgres
    // "ON CONFLICT DO UPDATE cannot affect row a second time" si el Excel repite un código.
    const byCode = new Map();
    for (const r of raw) {
      const m = mapVentaRow(r);
      if (m) byCode.set(m[0], m);
    }
    const rows = Array.from(byCode.values());
    if (rows.length === 0) {
      return res.status(400).json({ success: false, message: 'El archivo no tiene filas válidas (falta la columna CodigoCV).' });
    }

    await ensureVentasSchema();
    const client = await pgPool.connect();
    let insertados = 0, actualizados = 0;
    try {
      await client.query('BEGIN');
      const CHUNK = 1000;
      for (let i = 0; i < rows.length; i += CHUNK) {
        const r = await upsertVentasChunk(client, rows.slice(i, i + CHUNK));
        insertados += r.inserted;
        actualizados += r.updated;
      }
      await client.query(
        `INSERT INTO ventas_cargas (cargado_por, archivo, filas, insertados, actualizados)
         VALUES ($1,$2,$3,$4,$5)`,
        [toStr(req.body && req.body.cargado_por), req.file.originalname || null, rows.length, insertados, actualizados]
      );
      await client.query('COMMIT');
    } catch (e) {
      await client.query('ROLLBACK');
      throw e;
    } finally {
      client.release();
    }

    res.json({ success: true, filas: rows.length, insertados, actualizados, updated_at: new Date().toISOString() });
  } catch (error) {
    console.error('❌ Error en POST /ventas/import:', error);
    res.status(500).json({ success: false, message: 'No se pudo importar el archivo de ventas.' });
  }
});

// GET /ventas/estado — total de filas y datos de la última carga (para el "Actualizado al…").
app.get('/ventas/estado', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  try {
    await ensureVentasSchema();
    const { rows } = await pgPool.query('SELECT COUNT(*)::int AS total, MAX(updated_at) AS updated_at FROM ventas');
    const { rows: cargas } = await pgPool.query(
      'SELECT cargado_por, archivo, filas, insertados, actualizados, creado_en FROM ventas_cargas ORDER BY id DESC LIMIT 1'
    );
    res.json({ success: true, total: rows[0].total, updated_at: rows[0].updated_at, ultimaCarga: cargas[0] || null });
  } catch (error) {
    console.error('❌ Error en GET /ventas/estado:', error);
    res.status(500).json({ success: false, message: 'No se pudo obtener el estado de ventas.' });
  }
});

// GET /ventas?anio=&mes=&sede= — filas para los consumidores (ventas-sedes, pizarra…).
app.get('/ventas', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  try {
    await ensureVentasSchema();
    const cond = [];
    const params = [];
    const anio = req.query.anio ? parseInt(req.query.anio, 10) : null;
    const mes  = req.query.mes  ? parseInt(req.query.mes, 10)  : null;
    if (anio && mes) {
      // Trae ventas del mes por su fecha de venta (CV) Y las afectaciones
      // (NC/INC) cuya fecha de afectación (AF) cae en ese mes, aunque su venta
      // se haya registrado en otro mes. Necesario para que las NC cuadren.
      params.push(anio); const pa = params.length;
      params.push(mes);  const pm = params.length;
      cond.push(`((anio_cv = $${pa} AND mes_cv = $${pm}) OR (anio_af = $${pa} AND mes_af = $${pm}))`);
    } else if (anio) {
      params.push(anio);
      cond.push(`anio_cv = $${params.length}`);
    }
    if (req.query.sede) { params.push(`%${String(req.query.sede)}%`); cond.push(`sede ILIKE $${params.length}`); }
    const where = cond.length ? 'WHERE ' + cond.join(' AND ') : '';
    const { rows } = await pgPool.query(
      `SELECT * FROM ventas ${where} ORDER BY fecha_cv DESC NULLS LAST, codigo_cv DESC`, params
    );
    res.json(rows);
  } catch (error) {
    console.error('❌ Error en GET /ventas:', error);
    res.status(500).json({ success: false, message: 'No se pudieron obtener las ventas.' });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// 📊 MARGEN DE VENTAS → PostgreSQL (Neon)
//
// Fuente: Excel de márgenes a nivel de línea de producto. Un mismo CodigoCV puede
// tener VARIAS filas (una por producto), así que CodigoCV NO es único → no hay
// upsert por clave. La carga usa "reemplazo por CodigoCV": en una transacción se
// borran las filas de los CodigoCV presentes en el archivo y se reinsertan todas.
// Idempotente y correcto para el detalle uno-a-muchos.
//
// Endpoints:
//   POST /margen-ventas/import   (multipart, campo "archivo")
//   GET  /margen-ventas/estado
//   GET  /margen-ventas?anio=&mes=&sede=
// ─────────────────────────────────────────────────────────────────────────────

const MARGEN_COLS = [
  'codigo_cv', 'fecha', 'cliente', 'producto', 'marca', 'linea_producto',
  'cantidad', 'sede', 'linea_real', 'valor_venta', 'margen_total',
];

let margenSchemaLista = false;
async function ensureMargenSchema() {
  if (!pgPool || margenSchemaLista) return;
  await pgPool.query(`
    CREATE TABLE IF NOT EXISTS margen_ventas (
      id             BIGSERIAL PRIMARY KEY,
      codigo_cv      BIGINT,
      fecha          DATE,
      cliente        TEXT,
      producto       TEXT,
      marca          TEXT,
      linea_producto TEXT,
      cantidad       NUMERIC(14,2),
      sede           TEXT,
      linea_real     TEXT,
      valor_venta    NUMERIC(14,2),
      margen_total   NUMERIC(14,2),
      updated_at     TIMESTAMPTZ NOT NULL DEFAULT now(),
      created_at     TIMESTAMPTZ NOT NULL DEFAULT now()
    );
    CREATE INDEX IF NOT EXISTS ix_margen_codigo ON margen_ventas (codigo_cv);
    CREATE INDEX IF NOT EXISTS ix_margen_sede   ON margen_ventas (sede);
    CREATE INDEX IF NOT EXISTS ix_margen_fecha  ON margen_ventas (fecha);
    CREATE TABLE IF NOT EXISTS margen_ventas_cargas (
      id           BIGSERIAL PRIMARY KEY,
      cargado_por  TEXT,
      archivo      TEXT,
      filas        INTEGER,
      codigos      INTEGER,
      reemplazados INTEGER,
      creado_en    TIMESTAMPTZ NOT NULL DEFAULT now()
    );
  `);
  margenSchemaLista = true;
}

// Normaliza la columna Fecha a 'YYYY-MM-DD' (acepta Date, serial de Excel, dd/mm/yyyy o ISO).
function toFechaISO(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date && !isNaN(v)) {
    return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, '0')}-${String(v.getDate()).padStart(2, '0')}`;
  }
  if (typeof v === 'number') {
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    if (!isNaN(d)) return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
    return null;
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return `${m[3]}-${String(m[2]).padStart(2, '0')}-${String(m[1]).padStart(2, '0')}`;
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${String(m[2]).padStart(2, '0')}-${String(m[3]).padStart(2, '0')}`;
  const d = new Date(s);
  if (!isNaN(d)) return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  return null;
}

function mapMargenRow(r) {
  const codigo = toInt(pickCol(r, 'CodigoCV', 'codigo_cv', 'CODIGOCV', 'Codigo CV'));
  if (codigo === null) return null;
  return [
    codigo,
    toFechaISO(pickCol(r, 'Fecha', 'fecha', 'FECHA')),
    toStr(pickCol(r, 'Cliente', 'cliente')),
    toStr(pickCol(r, 'Producto', 'producto')),
    toStr(pickCol(r, 'Marca', 'marca')),
    toStr(pickCol(r, 'LineaProducto', 'Linea Producto', 'linea_producto')),
    toNum(pickCol(r, 'Cantidad', 'cantidad')),
    toStr(pickCol(r, 'SEDE', 'Sede', 'sede')),
    toStr(pickCol(r, 'LINEA REAL', 'LineaReal', 'linea_real')),
    toNum(pickCol(r, 'VALOR VENTA', 'ValorVenta', 'valor_venta')),
    toNum(pickCol(r, 'MARGEN TOTAL', 'MargenTotal', 'margen_total')),
  ];
}

async function insertMargenChunk(client, chunk) {
  const params = [];
  const tuples = chunk.map((row, i) => {
    const base = i * MARGEN_COLS.length;
    params.push(...row);
    return '(' + MARGEN_COLS.map((_, j) => `$${base + j + 1}`).join(',') + ')';
  });
  await client.query(`INSERT INTO margen_ventas (${MARGEN_COLS.join(',')}) VALUES ${tuples.join(',')}`, params);
}

app.post('/margen-ventas/import', upload.single('archivo'), async (req, res) => {
  if (!pgPool) {
    return res.status(500).json({ success: false, message: 'Base de datos no configurada (falta DATABASE_URL).' });
  }
  if (!req.file) {
    return res.status(400).json({ success: false, message: 'No se recibió archivo (campo "archivo").' });
  }

  try {
    const wb = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, { defval: '' });

    const rows = [];
    for (const r of raw) {
      const m = mapMargenRow(r);
      if (m) rows.push(m);
    }
    if (rows.length === 0) {
      return res.status(400).json({ success: false, message: 'El archivo no tiene filas válidas (falta la columna CodigoCV).' });
    }

    const codigos = Array.from(new Set(rows.map(r => r[0])));

    await ensureMargenSchema();
    const client = await pgPool.connect();
    let reemplazados = 0;
    try {
      await client.query('BEGIN');
      // Reemplazo por CodigoCV: borra los códigos presentes en el archivo…
      for (let i = 0; i < codigos.length; i += 5000) {
        const slice = codigos.slice(i, i + 5000);
        const del = await client.query('DELETE FROM margen_ventas WHERE codigo_cv = ANY($1::bigint[])', [slice]);
        reemplazados += del.rowCount;
      }
      // …y reinserta todas las filas del archivo.
      const CHUNK = 800;
      for (let i = 0; i < rows.length; i += CHUNK) {
        await insertMargenChunk(client, rows.slice(i, i + CHUNK));
      }
      await client.query(
        `INSERT INTO margen_ventas_cargas (cargado_por, archivo, filas, codigos, reemplazados)
         VALUES ($1,$2,$3,$4,$5)`,
        [toStr(req.body && req.body.cargado_por), req.file.originalname || null, rows.length, codigos.length, reemplazados]
      );
      await client.query('COMMIT');
    } catch (e) {
      await client.query('ROLLBACK');
      throw e;
    } finally {
      client.release();
    }

    res.json({ success: true, filas: rows.length, codigos: codigos.length, reemplazados, updated_at: new Date().toISOString() });
  } catch (error) {
    console.error('❌ Error en POST /margen-ventas/import:', error);
    res.status(500).json({ success: false, message: 'No se pudo importar el archivo de margen.' });
  }
});

app.get('/margen-ventas/estado', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  try {
    await ensureMargenSchema();
    const { rows } = await pgPool.query('SELECT COUNT(*)::int AS total, MAX(updated_at) AS updated_at FROM margen_ventas');
    const { rows: cargas } = await pgPool.query(
      'SELECT cargado_por, archivo, filas, codigos, reemplazados, creado_en FROM margen_ventas_cargas ORDER BY id DESC LIMIT 1'
    );
    res.json({ success: true, total: rows[0].total, updated_at: rows[0].updated_at, ultimaCarga: cargas[0] || null });
  } catch (error) {
    console.error('❌ Error en GET /margen-ventas/estado:', error);
    res.status(500).json({ success: false, message: 'No se pudo obtener el estado de margen.' });
  }
});

app.get('/margen-ventas', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  try {
    await ensureMargenSchema();
    const cond = [];
    const params = [];
    if (req.query.anio) { params.push(parseInt(req.query.anio, 10)); cond.push(`EXTRACT(YEAR FROM fecha) = $${params.length}`); }
    if (req.query.mes)  { params.push(parseInt(req.query.mes, 10));  cond.push(`EXTRACT(MONTH FROM fecha) = $${params.length}`); }
    if (req.query.sede) { params.push(`%${String(req.query.sede)}%`); cond.push(`sede ILIKE $${params.length}`); }
    const where = cond.length ? 'WHERE ' + cond.join(' AND ') : '';
    const { rows } = await pgPool.query(
      `SELECT * FROM margen_ventas ${where} ORDER BY fecha DESC NULLS LAST, codigo_cv DESC`, params
    );
    res.json(rows);
  } catch (error) {
    console.error('❌ Error en GET /margen-ventas:', error);
    res.status(500).json({ success: false, message: 'No se pudieron obtener los márgenes.' });
  }
});

// Iniciar servidor
app.listen(PORT, () => {
  console.log(`✅ API corriendo en http://localhost:${PORT}`);
  Promise.all([ensureVentasSchema(), ensureMargenSchema()])
    .then(() => pgPool && console.log('🐘 Esquemas de ventas y margen verificados.'))
    .catch((e) => console.error('❌ No se pudo verificar el esquema:', e));
});

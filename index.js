require('dotenv').config();
const express = require('express');
const { google } = require('googleapis');
const cors = require('cors');
const { Pool } = require('pg');
const multer = require('multer');
const XLSX = require('xlsx');
const bcrypt = require('bcryptjs');

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

// 🔐 Login: POST /auth/login — valida SOLO contra la BD (usuarios), bcrypt.
// El sheet ya no se usa en runtime (solo sirvió para la migración inicial).
app.post('/auth/login', async (req, res) => {
  const { usuario, password } = req.body;

  if (!usuario || !password) {
    return res.status(400).json({ success: false, message: 'Usuario y contraseña requeridos.' });
  }
  if (!pgPool) {
    return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  }

  try {
    await ensureUsuariosSchema();
    const { rows } = await pgPool.query('SELECT * FROM usuarios WHERE lower(usuario) = lower($1)', [usuario]);
    const u = rows[0];
    const ok = u && u.activo && await bcrypt.compare(password, u.password_hash || '');
    if (!ok) {
      return res.status(401).json({ success: false, message: 'Credenciales inválidas o usuario inactivo.' });
    }
    res.json({ success: true, nombre: u.nombre || '', rol: u.rol || '', sede: u.sede || '' });
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

// ─────────────────────────────────────────────────────────────────────────────
// 👤 USUARIOS → PostgreSQL (Neon). Contraseñas hasheadas con bcrypt.
// Reemplaza (con fallback) al sheet 'usuarios'. CRUD para el módulo Seguridad.
// ─────────────────────────────────────────────────────────────────────────────
let usuariosSchemaLista = false;
async function ensureUsuariosSchema() {
  if (!pgPool || usuariosSchemaLista) return;
  await pgPool.query(`
    CREATE TABLE IF NOT EXISTS usuarios (
      id             BIGSERIAL PRIMARY KEY,
      usuario        TEXT UNIQUE NOT NULL,
      password_hash  TEXT NOT NULL,
      nombre         TEXT,
      rol            TEXT,
      sede           TEXT,
      activo         BOOLEAN NOT NULL DEFAULT true,
      creado_en      TIMESTAMPTZ NOT NULL DEFAULT now(),
      actualizado_en TIMESTAMPTZ NOT NULL DEFAULT now()
    );
  `);
  usuariosSchemaLista = true;
}

// Migración única: si la tabla está vacía, importa los usuarios del sheet (hasheando).
async function migrarUsuariosDesdeSheet() {
  if (!pgPool) return;
  await ensureUsuariosSchema();
  const { rows } = await pgPool.query('SELECT COUNT(*)::int AS n FROM usuarios');
  if (rows[0].n > 0) return;
  try {
    const config = sheetsConfigs['usuarios'];
    const auth = googleAuthConfigs[config.authKey];
    const client = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: client });
    const resp = await sheets.spreadsheets.values.get({ spreadsheetId: config.spreadsheetId, range: config.range });
    const data = resp.data.values;
    if (!data || data.length < 2) return;
    const [headers, ...filas] = data;
    const idx = (c) => headers.indexOf(c);
    let n = 0;
    for (const row of filas) {
      const usuario = (row[idx('usuario')] || '').toString().trim();
      const pass = (row[idx('password')] || '').toString();
      if (!usuario || !pass) continue;
      const hash = await bcrypt.hash(pass, 10);
      const activo = (row[idx('activo')] || '').toString().trim().toUpperCase() === 'SI';
      await pgPool.query(
        `INSERT INTO usuarios (usuario, password_hash, nombre, rol, sede, activo)
         VALUES ($1,$2,$3,$4,$5,$6) ON CONFLICT (usuario) DO NOTHING`,
        [usuario, hash, (row[idx('nombre')] || '').toString().trim(),
         (row[idx('rol')] || '').toString().trim(), (row[idx('sede')] || '').toString().trim(), activo]
      );
      n++;
    }
    console.log(`🔐 Migrados ${n} usuarios del sheet a la BD.`);
  } catch (e) {
    console.error('⚠️ No se pudo migrar usuarios del sheet:', e.message);
  }
}

// GET /usuarios — lista (sin el hash).
app.get('/usuarios', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  try {
    await ensureUsuariosSchema();
    const { rows } = await pgPool.query(
      'SELECT id, usuario, nombre, rol, sede, activo, creado_en, actualizado_en FROM usuarios ORDER BY usuario'
    );
    res.json(rows);
  } catch (e) { console.error('❌ GET /usuarios', e); res.status(500).json({ success: false, message: 'No se pudieron obtener los usuarios.' }); }
});

// POST /usuarios — crea (hashea la contraseña).
app.post('/usuarios', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  const b = req.body || {};
  const usuario = (b.usuario || '').toString().trim();
  const password = (b.password || '').toString();
  if (!usuario || !password) return res.status(400).json({ success: false, message: 'Usuario y contraseña son obligatorios.' });
  try {
    await ensureUsuariosSchema();
    const hash = await bcrypt.hash(password, 10);
    const { rows } = await pgPool.query(
      `INSERT INTO usuarios (usuario, password_hash, nombre, rol, sede, activo)
       VALUES ($1,$2,$3,$4,$5,$6)
       RETURNING id, usuario, nombre, rol, sede, activo`,
      [usuario, hash, (b.nombre || '').toString().trim(), (b.rol || '').toString().trim(),
       (b.sede || '').toString().trim(), b.activo !== false]
    );
    res.json({ success: true, usuario: rows[0] });
  } catch (e) {
    if (e.code === '23505') return res.status(409).json({ success: false, message: 'Ya existe un usuario con ese nombre de acceso.' });
    console.error('❌ POST /usuarios', e); res.status(500).json({ success: false, message: 'No se pudo crear el usuario.' });
  }
});

// PUT /usuarios/:id — edita; si mandan password (no vacío) se rehashea.
app.put('/usuarios/:id', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  const b = req.body || {};
  const id = parseInt(req.params.id, 10);
  const usuario = (b.usuario || '').toString().trim();
  if (!usuario) return res.status(400).json({ success: false, message: 'El usuario es obligatorio.' });
  try {
    await ensureUsuariosSchema();
    const campos = ['usuario = $2', 'nombre = $3', 'rol = $4', 'sede = $5', 'activo = $6', 'actualizado_en = now()'];
    const params = [id, usuario, (b.nombre || '').toString().trim(), (b.rol || '').toString().trim(),
                    (b.sede || '').toString().trim(), b.activo !== false];
    if (b.password && b.password.toString().trim() !== '') {
      const hash = await bcrypt.hash(b.password.toString(), 10);
      params.push(hash);
      campos.push(`password_hash = $${params.length}`);
    }
    const { rows } = await pgPool.query(
      `UPDATE usuarios SET ${campos.join(', ')} WHERE id = $1
       RETURNING id, usuario, nombre, rol, sede, activo`, params
    );
    if (!rows.length) return res.status(404).json({ success: false, message: 'Usuario no encontrado.' });
    res.json({ success: true, usuario: rows[0] });
  } catch (e) {
    if (e.code === '23505') return res.status(409).json({ success: false, message: 'Ya existe un usuario con ese nombre de acceso.' });
    console.error('❌ PUT /usuarios/:id', e); res.status(500).json({ success: false, message: 'No se pudo actualizar el usuario.' });
  }
});

// PATCH /usuarios/:id/estado — activar / desactivar.
app.patch('/usuarios/:id/estado', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  const id = parseInt(req.params.id, 10);
  const activo = !!(req.body && req.body.activo);
  try {
    await ensureUsuariosSchema();
    const { rows } = await pgPool.query(
      'UPDATE usuarios SET activo = $2, actualizado_en = now() WHERE id = $1 RETURNING id, usuario, activo', [id, activo]
    );
    if (!rows.length) return res.status(404).json({ success: false, message: 'Usuario no encontrado.' });
    res.json({ success: true, usuario: rows[0] });
  } catch (e) { console.error('❌ PATCH /usuarios/:id/estado', e); res.status(500).json({ success: false, message: 'No se pudo cambiar el estado.' }); }
});

// ─────────────────────────────────────────────────────────────────────────────
// 🔑 PERMISOS (matriz Rol+Perfil → módulos) → PostgreSQL (Neon).
// Reemplaza el localStorage del navegador para que Seguridad sea centralizada.
// ─────────────────────────────────────────────────────────────────────────────
let permisosSchemaLista = false;
async function ensurePermisosSchema() {
  if (!pgPool || permisosSchemaLista) return;
  await pgPool.query(`
    CREATE TABLE IF NOT EXISTS permisos (
      clave          TEXT PRIMARY KEY,
      modulos        JSONB NOT NULL DEFAULT '[]'::jsonb,
      actualizado_en TIMESTAMPTZ NOT NULL DEFAULT now()
    );
  `);
  permisosSchemaLista = true;
}

// GET /permisos → { 'gerente-call': [...módulos], ... }
app.get('/permisos', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  try {
    await ensurePermisosSchema();
    const { rows } = await pgPool.query('SELECT clave, modulos FROM permisos');
    const map = {};
    rows.forEach(r => { map[r.clave] = r.modulos || []; });
    res.json(map);
  } catch (e) { console.error('❌ GET /permisos', e); res.status(500).json({ success: false, message: 'No se pudieron obtener los permisos.' }); }
});

// PUT /permisos → body = { clave: [módulos], ... } (upsert de todas las claves).
app.put('/permisos', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  const map = req.body || {};
  try {
    await ensurePermisosSchema();
    const client = await pgPool.connect();
    try {
      await client.query('BEGIN');
      for (const [clave, modulos] of Object.entries(map)) {
        await client.query(
          `INSERT INTO permisos (clave, modulos, actualizado_en) VALUES ($1, $2::jsonb, now())
           ON CONFLICT (clave) DO UPDATE SET modulos = EXCLUDED.modulos, actualizado_en = now()`,
          [clave, JSON.stringify(Array.isArray(modulos) ? modulos : [])]
        );
      }
      await client.query('COMMIT');
    } catch (e) { await client.query('ROLLBACK'); throw e; }
    finally { client.release(); }
    res.json({ success: true });
  } catch (e) { console.error('❌ PUT /permisos', e); res.status(500).json({ success: false, message: 'No se pudieron guardar los permisos.' }); }
});

// ─────────────────────────────────────────────────────────────────────────────
// 📋 GESTIÓN REALZZA → PostgreSQL (Neon). Reemplaza el Google Form de campo.
// La tabla guarda las 29 columnas del form + marca_temporal (real) + origen.
// GET devuelve las MISMAS cabeceras de la hoja para que los módulos que hoy
// consumen /data/campo funcionen igual cambiando una sola línea.
// ─────────────────────────────────────────────────────────────────────────────

// Columnas de la tabla, en el orden del INSERT.
const GRZ_COLS = [
  'marca_temporal', 'marca_temporal_raw', 'asesor_realzza', 'sede', 'tipo_base',
  'dni_cliente', 'celular_gestionado', 'estado_gestion', 'medio_primer_contacto',
  'resultado_gestion', 'producto_interes', 'motivo_interes', 'motivo_agendamiento',
  'fecha_interes_agendamiento', 'hora_interes_agendamiento', 'comentario_agendamiento',
  'fecha_interes_derivacion', 'hora_interes_derivacion', 'comentario_derivacion',
  'motivo_no_interes', 'comentario_no_interes', 'motivo_no_atendible', 'comentario_no_atendible',
  'motivos_tercero_relacionado', 'fecha_rellamada', 'hora_rellamada', 'numero_titular_actual',
  'motivo_no_contacto', 'motivo_no_cierre', 'comentario_venta_no_concretada', 'origen',
];

let grzSchemaLista = false;
async function ensureGestionRealzzaSchema() {
  if (!pgPool || grzSchemaLista) return;
  await pgPool.query(`
    CREATE TABLE IF NOT EXISTS gestion_realzza (
      id                            BIGSERIAL PRIMARY KEY,
      marca_temporal                TIMESTAMP,
      marca_temporal_raw            TEXT,
      asesor_realzza                TEXT,
      sede                          TEXT,
      tipo_base                     TEXT,
      dni_cliente                   TEXT,
      celular_gestionado            TEXT,
      estado_gestion                TEXT,
      medio_primer_contacto         TEXT,
      resultado_gestion             TEXT,
      producto_interes              TEXT,
      motivo_interes                TEXT,
      motivo_agendamiento           TEXT,
      fecha_interes_agendamiento    TEXT,
      hora_interes_agendamiento     TEXT,
      comentario_agendamiento       TEXT,
      fecha_interes_derivacion      TEXT,
      hora_interes_derivacion       TEXT,
      comentario_derivacion         TEXT,
      motivo_no_interes             TEXT,
      comentario_no_interes         TEXT,
      motivo_no_atendible           TEXT,
      comentario_no_atendible       TEXT,
      motivos_tercero_relacionado   TEXT,
      fecha_rellamada               TEXT,
      hora_rellamada                TEXT,
      numero_titular_actual         TEXT,
      motivo_no_contacto            TEXT,
      motivo_no_cierre              TEXT,
      comentario_venta_no_concretada TEXT,
      origen                        TEXT NOT NULL DEFAULT 'app',
      creado_en                     TIMESTAMPTZ NOT NULL DEFAULT now()
    );
    CREATE INDEX IF NOT EXISTS ix_grz_marca  ON gestion_realzza (marca_temporal);
    CREATE INDEX IF NOT EXISTS ix_grz_asesor ON gestion_realzza (asesor_realzza);
    CREATE INDEX IF NOT EXISTS ix_grz_dni    ON gestion_realzza (dni_cliente);
  `);
  grzSchemaLista = true;
}

// "14/10/2025 9:18:07" → Date. Devuelve null si no parsea.
function parseMarcaTemporal(s) {
  if (!s) return null;
  const [fecha, hora] = s.toString().trim().split(' ');
  const [d, m, y] = (fecha || '').split('/').map(Number);
  if (!d || !m || !y) return null;
  const [hh = 0, mm = 0, ss = 0] = (hora || '').split(':').map(Number);
  const dt = new Date(y, m - 1, d, hh || 0, mm || 0, ss || 0);
  return isNaN(dt) ? null : dt;
}
function formatMarca(d) {
  const p = n => String(n).padStart(2, '0');
  return `${d.getDate()}/${d.getMonth() + 1}/${d.getFullYear()} ${d.getHours()}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}

// Fecha/hora actual en Perú (America/Lima, UTC-5) SIN depender de la zona del
// servidor (en Render es UTC → +5h). Devuelve:
//  - ts:  'yyyy-mm-dd HH:mm:ss'  para la columna TIMESTAMP (se guarda tal cual).
//  - raw: 'd/m/yyyy H:mm:ss'     mismo formato que formatMarca (marca_temporal_raw).
function ahoraLima() {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'America/Lima',
    year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false,
  }).formatToParts(new Date());
  const p = {};
  for (const x of parts) if (x.type !== 'literal') p[x.type] = x.value;
  const hh = p.hour === '24' ? '00' : p.hour;   // Intl a veces da '24' a medianoche
  return {
    ts: `${p.year}-${p.month}-${p.day} ${hh}:${p.minute}:${p.second}`,
    raw: `${+p.day}/${+p.month}/${+p.year} ${+hh}:${p.minute}:${p.second}`,
  };
}

// Fila de la hoja (objeto por cabecera) → arreglo en el orden de GRZ_COLS.
function mapRealzzaRow(r, origen) {
  return [
    parseMarcaTemporal(r['Marca temporal']),
    toStr(r['Marca temporal']),
    toStr(r['ASESOR REALZZA']), toStr(r['SEDE']), toStr(r['TIPO DE BASE']),
    toStr(r['DNI CLIENTE']), toStr(r['CELULAR GESTIONADO']), toStr(r['ESTADO DE GESTIÓN']),
    toStr(r['MEDIO DE PRIMER CONTACTO']), toStr(r['RESULTADO DE GESTIÓN']), toStr(r['PRODUCTO INTERÉS']),
    toStr(r['MOTIVO INTERÉS']), toStr(r['MOTIVO AGENDAMIENTO']), toStr(r['FECHA DE INTERÉS AGENDAMIENTO']),
    toStr(r['HORA APROXIMADA INTERÉS AGENDAMIENTO']), toStr(r['COMENTARIO ADICIONAL AGENDAMIENTO']),
    toStr(r['FECHA DE INTERÉS DERIVACIÓN']), toStr(r['HORA APROXIMADA INTERÉS DERIVACIÓN']),
    toStr(r['COMENTARIO ADICIONAL DERIVACIÓN']), toStr(r['MOTIVO NO INTERÉS']), toStr(r['COMENTARIO ADICIONAL NO INTERÉS']),
    toStr(r['MOTIVO NO ATENDIBLE']), toStr(r['COMENTARIO ADICIONAL NO ATENDIBLE']), toStr(r['MOTIVOS TERCERO RELACIONADO']),
    toStr(r['FECHA DE RE-LLAMADA']), toStr(r['HORA DE RELLAMADA']), toStr(r['NÚMERO TITULAR ACTUAL']),
    toStr(r['MOTIVO NO CONTACTO']), toStr(r['MOTIVO DE NO CIERRE']), toStr(r['COMENTARIO VENTA NO CONCRETADA']),
    origen,
  ];
}

// Fila de la BD → objeto con las MISMAS cabeceras de la hoja (para los consumidores).
function grzRowToSheet(row) {
  return {
    id: row.id,
    'Marca temporal': row.marca_temporal_raw || '',
    'ASESOR REALZZA': row.asesor_realzza || '',
    'SEDE': row.sede || '',
    'TIPO DE BASE': row.tipo_base || '',
    'DNI CLIENTE': row.dni_cliente || '',
    'CELULAR GESTIONADO': row.celular_gestionado || '',
    'ESTADO DE GESTIÓN': row.estado_gestion || '',
    'MEDIO DE PRIMER CONTACTO': row.medio_primer_contacto || '',
    'RESULTADO DE GESTIÓN': row.resultado_gestion || '',
    'PRODUCTO INTERÉS': row.producto_interes || '',
    'MOTIVO INTERÉS': row.motivo_interes || '',
    'MOTIVO AGENDAMIENTO': row.motivo_agendamiento || '',
    'FECHA DE INTERÉS AGENDAMIENTO': row.fecha_interes_agendamiento || '',
    'HORA APROXIMADA INTERÉS AGENDAMIENTO': row.hora_interes_agendamiento || '',
    'COMENTARIO ADICIONAL AGENDAMIENTO': row.comentario_agendamiento || '',
    'FECHA DE INTERÉS DERIVACIÓN': row.fecha_interes_derivacion || '',
    'HORA APROXIMADA INTERÉS DERIVACIÓN': row.hora_interes_derivacion || '',
    'COMENTARIO ADICIONAL DERIVACIÓN': row.comentario_derivacion || '',
    'MOTIVO NO INTERÉS': row.motivo_no_interes || '',
    'COMENTARIO ADICIONAL NO INTERÉS': row.comentario_no_interes || '',
    'MOTIVO NO ATENDIBLE': row.motivo_no_atendible || '',
    'COMENTARIO ADICIONAL NO ATENDIBLE': row.comentario_no_atendible || '',
    'MOTIVOS TERCERO RELACIONADO': row.motivos_tercero_relacionado || '',
    'FECHA DE RE-LLAMADA': row.fecha_rellamada || '',
    'HORA DE RELLAMADA': row.hora_rellamada || '',
    'NÚMERO TITULAR ACTUAL': row.numero_titular_actual || '',
    'MOTIVO NO CONTACTO': row.motivo_no_contacto || '',
    'MOTIVO DE NO CIERRE': row.motivo_no_cierre || '',
    'COMENTARIO VENTA NO CONCRETADA': row.comentario_venta_no_concretada || '',
  };
}

// Lee la hoja de respuestas del form Realzza (campo) como objetos por cabecera.
async function leerCampoSheet() {
  const config = sheetsConfigs['campo'];
  const auth = googleAuthConfigs[config.authKey];
  const client = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });
  const resp = await sheets.spreadsheets.values.get({ spreadsheetId: config.spreadsheetId, range: config.range });
  const rows = resp.data.values || [];
  if (rows.length < 2) return [];
  const [headersRaw, ...data] = rows;
  const seen = {}; const H = [];
  headersRaw.forEach(h => { if (!seen[h]) { seen[h] = 1; H.push(h); } else { H.push(`${h} (${seen[h]})`); seen[h]++; } });
  return data.map(row => H.reduce((acc, h, i) => { acc[h] = row[i] || ''; return acc; }, {}));
}

// POST /gestion-realzza/sync — migra una sola vez las respuestas del form a la BD.
app.post('/gestion-realzza/sync', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  try {
    await ensureGestionRealzzaSchema();
    const { rows: cnt } = await pgPool.query("SELECT COUNT(*)::int AS n FROM gestion_realzza WHERE origen = 'form'");
    if (cnt[0].n > 0 && req.query.force !== '1') {
      return res.json({ success: true, yaMigrado: true, existentes: cnt[0].n });
    }
    const data = await leerCampoSheet();
    const filas = data.filter(r =>
      (r['Marca temporal'] || '').toString().trim() !== '' || (r['ASESOR REALZZA'] || '').toString().trim() !== '');

    const client = await pgPool.connect();
    let insertados = 0;
    try {
      await client.query('BEGIN');
      const CHUNK = 500;
      for (let i = 0; i < filas.length; i += CHUNK) {
        const chunk = filas.slice(i, i + CHUNK);
        const params = [];
        const tuples = chunk.map((r, idx) => {
          const arr = mapRealzzaRow(r, 'form');
          const base = idx * GRZ_COLS.length;
          params.push(...arr);
          return '(' + GRZ_COLS.map((_, j) => `$${base + j + 1}`).join(',') + ')';
        });
        await client.query(`INSERT INTO gestion_realzza (${GRZ_COLS.join(',')}) VALUES ${tuples.join(',')}`, params);
        insertados += chunk.length;
      }
      await client.query('COMMIT');
    } catch (e) { await client.query('ROLLBACK'); throw e; } finally { client.release(); }

    res.json({ success: true, leidas: data.length, insertados });
  } catch (e) {
    console.error('❌ POST /gestion-realzza/sync:', e);
    res.status(500).json({ success: false, message: e.message });
  }
});

// POST /gestion-realzza — registra una gestión nueva desde la app (origen = 'app').
app.post('/gestion-realzza', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  const b = req.body || {};
  if (!b.asesor_realzza || !b.dni_cliente || !b.estado_gestion) {
    return res.status(400).json({ success: false, message: 'Faltan campos obligatorios (asesor, dni, estado de gestión).' });
  }
  try {
    await ensureGestionRealzzaSchema();
    const t = ahoraLima();
    const valorDe = {
      marca_temporal: t.ts, marca_temporal_raw: t.raw, origen: 'app',
      asesor_realzza: b.asesor_realzza, sede: b.sede || 'REALZZA', tipo_base: b.tipo_base,
      dni_cliente: b.dni_cliente, celular_gestionado: b.celular_gestionado, estado_gestion: b.estado_gestion,
      medio_primer_contacto: b.medio_primer_contacto, resultado_gestion: b.resultado_gestion,
      producto_interes: b.producto_interes, motivo_interes: b.motivo_interes, motivo_agendamiento: b.motivo_agendamiento,
      fecha_interes_agendamiento: b.fecha_interes_agendamiento, hora_interes_agendamiento: b.hora_interes_agendamiento,
      comentario_agendamiento: b.comentario_agendamiento, fecha_interes_derivacion: b.fecha_interes_derivacion,
      hora_interes_derivacion: b.hora_interes_derivacion, comentario_derivacion: b.comentario_derivacion,
      motivo_no_interes: b.motivo_no_interes, comentario_no_interes: b.comentario_no_interes,
      motivo_no_atendible: b.motivo_no_atendible, comentario_no_atendible: b.comentario_no_atendible,
      motivos_tercero_relacionado: b.motivos_tercero_relacionado, fecha_rellamada: b.fecha_rellamada,
      hora_rellamada: b.hora_rellamada, numero_titular_actual: b.numero_titular_actual,
      motivo_no_contacto: b.motivo_no_contacto, motivo_no_cierre: b.motivo_no_cierre,
      comentario_venta_no_concretada: b.comentario_venta_no_concretada,
    };
    const params = GRZ_COLS.map(c => {
      const v = valorDe[c];
      return v === undefined || v === '' ? null : v;
    });
    const ph = GRZ_COLS.map((_, i) => `$${i + 1}`).join(',');
    const { rows } = await pgPool.query(
      `INSERT INTO gestion_realzza (${GRZ_COLS.join(',')}) VALUES (${ph}) RETURNING id`, params);
    res.json({ success: true, id: rows[0].id, marca_temporal: valorDe.marca_temporal_raw });
  } catch (e) {
    console.error('❌ POST /gestion-realzza:', e);
    res.status(500).json({ success: false, message: 'No se pudo guardar la gestión.' });
  }
});

// GET /gestion-realzza?desde=&hasta= — filas con las cabeceras de la hoja.
app.get('/gestion-realzza', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  try {
    await ensureGestionRealzzaSchema();
    const cond = []; const params = [];
    if (req.query.desde) { params.push(`${req.query.desde} 00:00:00`); cond.push(`marca_temporal >= $${params.length}`); }
    if (req.query.hasta) { params.push(`${req.query.hasta} 23:59:59`); cond.push(`marca_temporal <= $${params.length}`); }
    const where = cond.length ? 'WHERE ' + cond.join(' AND ') : '';
    const { rows } = await pgPool.query(
      `SELECT * FROM gestion_realzza ${where} ORDER BY marca_temporal DESC NULLS LAST`, params);
    res.json(rows.map(grzRowToSheet));
  } catch (e) {
    console.error('❌ GET /gestion-realzza:', e);
    res.status(500).json({ success: false, message: 'No se pudieron obtener las gestiones.' });
  }
});

// Mapa cabecera de hoja → columna BD (Realzza), para editar desde el grid.
const REALZZA_SHEET_TO_COL = {
  'ASESOR REALZZA': 'asesor_realzza', 'SEDE': 'sede', 'TIPO DE BASE': 'tipo_base', 'DNI CLIENTE': 'dni_cliente',
  'CELULAR GESTIONADO': 'celular_gestionado', 'ESTADO DE GESTIÓN': 'estado_gestion', 'MEDIO DE PRIMER CONTACTO': 'medio_primer_contacto',
  'RESULTADO DE GESTIÓN': 'resultado_gestion', 'PRODUCTO INTERÉS': 'producto_interes', 'MOTIVO INTERÉS': 'motivo_interes',
  'MOTIVO AGENDAMIENTO': 'motivo_agendamiento', 'FECHA DE INTERÉS AGENDAMIENTO': 'fecha_interes_agendamiento',
  'HORA APROXIMADA INTERÉS AGENDAMIENTO': 'hora_interes_agendamiento', 'COMENTARIO ADICIONAL AGENDAMIENTO': 'comentario_agendamiento',
  'FECHA DE INTERÉS DERIVACIÓN': 'fecha_interes_derivacion', 'HORA APROXIMADA INTERÉS DERIVACIÓN': 'hora_interes_derivacion',
  'COMENTARIO ADICIONAL DERIVACIÓN': 'comentario_derivacion', 'MOTIVO NO INTERÉS': 'motivo_no_interes',
  'COMENTARIO ADICIONAL NO INTERÉS': 'comentario_no_interes', 'MOTIVO NO ATENDIBLE': 'motivo_no_atendible',
  'COMENTARIO ADICIONAL NO ATENDIBLE': 'comentario_no_atendible', 'MOTIVOS TERCERO RELACIONADO': 'motivos_tercero_relacionado',
  'FECHA DE RE-LLAMADA': 'fecha_rellamada', 'HORA DE RELLAMADA': 'hora_rellamada', 'NÚMERO TITULAR ACTUAL': 'numero_titular_actual',
  'MOTIVO NO CONTACTO': 'motivo_no_contacto',
  'MOTIVO DE NO CIERRE': 'motivo_no_cierre', 'COMENTARIO VENTA NO CONCRETADA': 'comentario_venta_no_concretada',
};

// Construye SET clause aceptando claves de hoja o snake_case (no toca marca_temporal/origen/id).
function construirUpdate(body, sheetToCol, colsSnake) {
  const sets = [], params = [];
  for (const [k, v] of Object.entries(body || {})) {
    const col = sheetToCol[k] || (colsSnake.includes(k) ? k : null);
    if (!col || col === 'marca_temporal' || col === 'marca_temporal_raw' || col === 'origen') continue;
    params.push(v === '' ? null : v);
    sets.push(`${col} = $${params.length}`);
  }
  return { sets, params };
}

// PUT /gestion-realzza/:id — edita una gestión.
app.put('/gestion-realzza/:id', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  try {
    await ensureGestionRealzzaSchema();
    const { sets, params } = construirUpdate(req.body, REALZZA_SHEET_TO_COL, GRZ_COLS);
    if (!sets.length) return res.status(400).json({ success: false, message: 'Nada para actualizar.' });
    params.push(parseInt(req.params.id, 10));
    const { rowCount } = await pgPool.query(`UPDATE gestion_realzza SET ${sets.join(', ')} WHERE id = $${params.length}`, params);
    if (!rowCount) return res.status(404).json({ success: false, message: 'Gestión no encontrada.' });
    res.json({ success: true });
  } catch (e) { console.error('❌ PUT /gestion-realzza/:id', e); res.status(500).json({ success: false, message: 'No se pudo actualizar.' }); }
});

// DELETE /gestion-realzza/:id — elimina una gestión.
app.delete('/gestion-realzza/:id', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  try {
    await ensureGestionRealzzaSchema();
    const { rowCount } = await pgPool.query('DELETE FROM gestion_realzza WHERE id = $1', [parseInt(req.params.id, 10)]);
    if (!rowCount) return res.status(404).json({ success: false, message: 'Gestión no encontrada.' });
    res.json({ success: true });
  } catch (e) { console.error('❌ DELETE /gestion-realzza/:id', e); res.status(500).json({ success: false, message: 'No se pudo eliminar.' }); }
});

// POST /gestion-realzza/match — { dnis: [...] } → { <dni>: {asesor, tipo_base, sede, celular} } (última gestión).
app.post('/gestion-realzza/match', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  try {
    await ensureGestionRealzzaSchema();
    const dnis = Array.from(new Set((req.body?.dnis || []).map(d => String(d).replace(/\D/g, '').replace(/^0+/, '')).filter(Boolean)));
    if (!dnis.length) return res.json({});
    const { rows } = await pgPool.query(`
      SELECT DISTINCT ON (dnin) dnin, asesor_realzza, tipo_base, sede, celular_gestionado
      FROM (
        SELECT regexp_replace(regexp_replace(dni_cliente, '\\D', '', 'g'), '^0+', '') AS dnin,
               asesor_realzza, tipo_base, sede, celular_gestionado, marca_temporal
        FROM gestion_realzza
        WHERE regexp_replace(regexp_replace(dni_cliente, '\\D', '', 'g'), '^0+', '') = ANY($1)
      ) t
      ORDER BY dnin, marca_temporal DESC NULLS LAST
    `, [dnis]);
    const map = {};
    rows.forEach(r => { map[r.dnin] = { asesor: r.asesor_realzza || '', tipo_base: r.tipo_base || '', sede: r.sede || '', celular: r.celular_gestionado || '' }; });
    res.json(map);
  } catch (e) { console.error('❌ POST /gestion-realzza/match', e); res.status(500).json({ success: false, message: 'No se pudo hacer el match.' }); }
});

// ─────────────────────────────────────────────────────────────────────────────
// 🕵️ CONTROL DEL SUPERVISOR (Realzza) → PostgreSQL (Neon).
// El supervisor registra su control de gestión: DNI, celular, estado (CONTACTO/NO
// CONTACTO), comentario + asesor y tipo de base. Se cruza luego con la gestión del
// asesor por DNI+celular para verificar si esa gestión fue supervisada.
// ─────────────────────────────────────────────────────────────────────────────
const CS_COLS = [
  'marca_temporal', 'marca_temporal_raw', 'registrado_por', 'asesor', 'tipo_base',
  'dni_cliente', 'celular', 'estado_gestion', 'comentario',
];

let csSchemaLista = false;
async function ensureControlSupervisorSchema() {
  if (!pgPool || csSchemaLista) return;
  await pgPool.query(`
    CREATE TABLE IF NOT EXISTS control_supervisor (
      id                 BIGSERIAL PRIMARY KEY,
      marca_temporal     TIMESTAMP,
      marca_temporal_raw TEXT,
      registrado_por     TEXT,
      asesor             TEXT,
      tipo_base          TEXT,
      dni_cliente        TEXT,
      celular            TEXT,
      estado_gestion     TEXT,
      comentario         TEXT,
      creado_en          TIMESTAMPTZ NOT NULL DEFAULT now()
    );
    CREATE INDEX IF NOT EXISTS ix_cs_marca  ON control_supervisor (marca_temporal);
    CREATE INDEX IF NOT EXISTS ix_cs_asesor ON control_supervisor (asesor);
    CREATE INDEX IF NOT EXISTS ix_cs_dni    ON control_supervisor (dni_cliente);
  `);
  csSchemaLista = true;
}

// Fila de la BD → objeto JSON para el frontend.
function csRowToJson(row) {
  return {
    id: row.id,
    marca_temporal: row.marca_temporal_raw || '',
    registrado_por: row.registrado_por || '',
    asesor: row.asesor || '',
    tipo_base: row.tipo_base || '',
    dni_cliente: row.dni_cliente || '',
    celular: row.celular || '',
    estado_gestion: row.estado_gestion || '',
    comentario: row.comentario || '',
  };
}

// POST /control-supervisor — registra un control del supervisor.
app.post('/control-supervisor', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  const b = req.body || {};
  if (!b.dni_cliente || !b.estado_gestion) {
    return res.status(400).json({ success: false, message: 'Faltan campos obligatorios (dni, estado de gestión).' });
  }
  try {
    await ensureControlSupervisorSchema();
    const t = ahoraLima();
    const valorDe = {
      marca_temporal: t.ts, marca_temporal_raw: t.raw,
      registrado_por: b.registrado_por, asesor: b.asesor, tipo_base: b.tipo_base,
      dni_cliente: b.dni_cliente, celular: b.celular, estado_gestion: b.estado_gestion,
      comentario: b.comentario,
    };
    const params = CS_COLS.map(c => { const v = valorDe[c]; return v === undefined || v === '' ? null : v; });
    const ph = CS_COLS.map((_, i) => `$${i + 1}`).join(',');
    const { rows } = await pgPool.query(
      `INSERT INTO control_supervisor (${CS_COLS.join(',')}) VALUES (${ph}) RETURNING id`, params);
    res.json({ success: true, id: rows[0].id, marca_temporal: valorDe.marca_temporal_raw });
  } catch (e) {
    console.error('❌ POST /control-supervisor:', e);
    res.status(500).json({ success: false, message: 'No se pudo guardar el control.' });
  }
});

// GET /control-supervisor?desde=&hasta= — controles del supervisor por rango de fechas.
app.get('/control-supervisor', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  try {
    await ensureControlSupervisorSchema();
    const cond = []; const params = [];
    if (req.query.desde) { params.push(`${req.query.desde} 00:00:00`); cond.push(`marca_temporal >= $${params.length}`); }
    if (req.query.hasta) { params.push(`${req.query.hasta} 23:59:59`); cond.push(`marca_temporal <= $${params.length}`); }
    const where = cond.length ? 'WHERE ' + cond.join(' AND ') : '';
    const { rows } = await pgPool.query(
      `SELECT * FROM control_supervisor ${where} ORDER BY marca_temporal DESC NULLS LAST`, params);
    res.json(rows.map(csRowToJson));
  } catch (e) {
    console.error('❌ GET /control-supervisor:', e);
    res.status(500).json({ success: false, message: 'No se pudieron obtener los controles.' });
  }
});

// Mapa clave → columna BD (para editar desde el grid; no toca marca_temporal/id).
const CS_SHEET_TO_COL = {
  registrado_por: 'registrado_por', asesor: 'asesor', tipo_base: 'tipo_base', dni_cliente: 'dni_cliente',
  celular: 'celular', estado_gestion: 'estado_gestion', comentario: 'comentario',
};

// PUT /control-supervisor/:id — edita un control.
app.put('/control-supervisor/:id', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  try {
    await ensureControlSupervisorSchema();
    const { sets, params } = construirUpdate(req.body, CS_SHEET_TO_COL, CS_COLS);
    if (!sets.length) return res.status(400).json({ success: false, message: 'Nada para actualizar.' });
    params.push(parseInt(req.params.id, 10));
    const { rowCount } = await pgPool.query(`UPDATE control_supervisor SET ${sets.join(', ')} WHERE id = $${params.length}`, params);
    if (!rowCount) return res.status(404).json({ success: false, message: 'Control no encontrado.' });
    res.json({ success: true });
  } catch (e) { console.error('❌ PUT /control-supervisor/:id', e); res.status(500).json({ success: false, message: 'No se pudo actualizar.' }); }
});

// DELETE /control-supervisor/:id — elimina un control.
app.delete('/control-supervisor/:id', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  try {
    await ensureControlSupervisorSchema();
    const { rowCount } = await pgPool.query('DELETE FROM control_supervisor WHERE id = $1', [parseInt(req.params.id, 10)]);
    if (!rowCount) return res.status(404).json({ success: false, message: 'Control no encontrado.' });
    res.json({ success: true });
  } catch (e) { console.error('❌ DELETE /control-supervisor/:id', e); res.status(500).json({ success: false, message: 'No se pudo eliminar.' }); }
});

// ─────────────────────────────────────────────────────────────────────────────
// 📞 GESTIÓN CALL CENTER → PostgreSQL (Neon). Reemplaza el Google Form de call.
// Misma mecánica que gestión realzza (30 columnas del form + marca_temporal + origen).
// GET devuelve las MISMAS cabeceras de la hoja /data/call.
// ─────────────────────────────────────────────────────────────────────────────
const GC_COLS = [
  'marca_temporal', 'marca_temporal_raw', 'asesor_contact', 'dni_cliente', 'tipo_cliente',
  'estado_gestion', 'medio_primer_contacto', 'celular_gestionado', 'resultado_gestion',
  'producto_interes', 'motivo_interes', 'motivo_agendamiento', 'fecha_interes_agendamiento',
  'hora_interes_agendamiento', 'fecha_interes_derivacion', 'hora_interes_derivacion',
  'comentario_derivacion', 'comentario_agendamiento', 'motivo_no_interes', 'comentario_no_interes',
  'motivo_no_atendible', 'comentario_no_atendible', 'motivos_tercero_relacionado', 'fecha_rellamada',
  'hora_rellamada', 'numero_titular_actual', 'motivo_no_contacto', 'sede', 'kommo',
  'motivo_no_cierre', 'comentario_venta_no_concretada', 'origen',
];

let gcSchemaLista = false;
async function ensureGestionCallSchema() {
  if (!pgPool || gcSchemaLista) return;
  await pgPool.query(`
    CREATE TABLE IF NOT EXISTS gestion_call (
      id                             BIGSERIAL PRIMARY KEY,
      marca_temporal                 TIMESTAMP,
      marca_temporal_raw             TEXT,
      asesor_contact                 TEXT,
      dni_cliente                    TEXT,
      tipo_cliente                   TEXT,
      estado_gestion                 TEXT,
      medio_primer_contacto          TEXT,
      celular_gestionado             TEXT,
      resultado_gestion              TEXT,
      producto_interes               TEXT,
      motivo_interes                 TEXT,
      motivo_agendamiento            TEXT,
      fecha_interes_agendamiento     TEXT,
      hora_interes_agendamiento      TEXT,
      fecha_interes_derivacion       TEXT,
      hora_interes_derivacion        TEXT,
      comentario_derivacion          TEXT,
      comentario_agendamiento        TEXT,
      motivo_no_interes              TEXT,
      comentario_no_interes          TEXT,
      motivo_no_atendible            TEXT,
      comentario_no_atendible        TEXT,
      motivos_tercero_relacionado    TEXT,
      fecha_rellamada                TEXT,
      hora_rellamada                 TEXT,
      numero_titular_actual          TEXT,
      motivo_no_contacto             TEXT,
      sede                           TEXT,
      kommo                          TEXT,
      motivo_no_cierre               TEXT,
      comentario_venta_no_concretada TEXT,
      origen                         TEXT NOT NULL DEFAULT 'app',
      creado_en                      TIMESTAMPTZ NOT NULL DEFAULT now()
    );
    CREATE INDEX IF NOT EXISTS ix_gc_marca  ON gestion_call (marca_temporal);
    CREATE INDEX IF NOT EXISTS ix_gc_asesor ON gestion_call (asesor_contact);
    CREATE INDEX IF NOT EXISTS ix_gc_dni    ON gestion_call (dni_cliente);
  `);
  gcSchemaLista = true;
}

function mapCallRow(r, origen) {
  return [
    parseMarcaTemporal(r['Marca temporal']), toStr(r['Marca temporal']),
    toStr(r['ASESOR CONTACT']), toStr(r['DNI CLIENTE']), toStr(r['TIPO DE CLIENTE']), toStr(r['ESTADO DE GESTIÓN']),
    toStr(r['MEDIO DE PRIMER CONTACTO']), toStr(r['CELULAR GESTIONADO']), toStr(r['RESULTADO DE GESTIÓN']),
    toStr(r['PRODUCTO INTERÉS']), toStr(r['MOTIVO INTERÉS']), toStr(r['MOTIVO AGENDAMIENTO']),
    toStr(r['FECHA DE INTERÉS AGENDAMIENTO']), toStr(r['HORA APROXIMADA INTERÉS AGENDAMIENTO']),
    toStr(r['FECHA DE INTERÉS DERIVACIÓN']), toStr(r['HORA APROXIMADA INTERÉS DERIVACIÓN']),
    toStr(r['COMENTARIO ADICIONAL DERIVACIÓN']), toStr(r['COMENTARIO ADICIONAL AGENDAMIENTO']),
    toStr(r['MOTIVO NO INTERÉS']), toStr(r['COMENTARIO ADICIONAL NO INTERES']),
    toStr(r['MOTIVO NO ATENDIBLE']), toStr(r['COMENTARIO ADICIONAL NO ATENDIBLE']),
    toStr(r['MOTIVOS TERCERO RELACIONADO']), toStr(r['FECHA DE RE-LLAMADA']), toStr(r['HORA DE RELLAMADA']),
    toStr(r['NÚMERO TITULAR ACTUAL']), toStr(r['MOTIVO NO CONTACTO']), toStr(r['SEDE']), toStr(r['KOMMO']),
    toStr(r['MOTIVO DE NO CIERRE']), toStr(r['COMENTARIO VENTA NO CONCRETADA']),
    origen,
  ];
}

function gcRowToSheet(row) {
  return {
    id: row.id,
    'Marca temporal': row.marca_temporal_raw || '',
    'ASESOR CONTACT': row.asesor_contact || '',
    'DNI CLIENTE': row.dni_cliente || '',
    'TIPO DE CLIENTE': row.tipo_cliente || '',
    'ESTADO DE GESTIÓN': row.estado_gestion || '',
    'MEDIO DE PRIMER CONTACTO': row.medio_primer_contacto || '',
    'CELULAR GESTIONADO': row.celular_gestionado || '',
    'RESULTADO DE GESTIÓN': row.resultado_gestion || '',
    'PRODUCTO INTERÉS': row.producto_interes || '',
    'MOTIVO INTERÉS': row.motivo_interes || '',
    'MOTIVO AGENDAMIENTO': row.motivo_agendamiento || '',
    'FECHA DE INTERÉS AGENDAMIENTO': row.fecha_interes_agendamiento || '',
    'HORA APROXIMADA INTERÉS AGENDAMIENTO': row.hora_interes_agendamiento || '',
    'FECHA DE INTERÉS DERIVACIÓN': row.fecha_interes_derivacion || '',
    'HORA APROXIMADA INTERÉS DERIVACIÓN': row.hora_interes_derivacion || '',
    'COMENTARIO ADICIONAL DERIVACIÓN': row.comentario_derivacion || '',
    'COMENTARIO ADICIONAL AGENDAMIENTO': row.comentario_agendamiento || '',
    'MOTIVO NO INTERÉS': row.motivo_no_interes || '',
    'COMENTARIO ADICIONAL NO INTERES': row.comentario_no_interes || '',
    'MOTIVO NO ATENDIBLE': row.motivo_no_atendible || '',
    'COMENTARIO ADICIONAL NO ATENDIBLE': row.comentario_no_atendible || '',
    'MOTIVOS TERCERO RELACIONADO': row.motivos_tercero_relacionado || '',
    'FECHA DE RE-LLAMADA': row.fecha_rellamada || '',
    'HORA DE RELLAMADA': row.hora_rellamada || '',
    'NÚMERO TITULAR ACTUAL': row.numero_titular_actual || '',
    'MOTIVO NO CONTACTO': row.motivo_no_contacto || '',
    'SEDE': row.sede || '',
    'KOMMO': row.kommo || '',
    'MOTIVO DE NO CIERRE': row.motivo_no_cierre || '',
    'COMENTARIO VENTA NO CONCRETADA': row.comentario_venta_no_concretada || '',
  };
}

async function leerCallSheet() {
  const config = sheetsConfigs['call'];
  const auth = googleAuthConfigs[config.authKey];
  const client = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });
  const resp = await sheets.spreadsheets.values.get({ spreadsheetId: config.spreadsheetId, range: config.range });
  const rows = resp.data.values || [];
  if (rows.length < 2) return [];
  const [headersRaw, ...data] = rows;
  const seen = {}; const H = [];
  headersRaw.forEach(h => { if (!seen[h]) { seen[h] = 1; H.push(h); } else { H.push(`${h} (${seen[h]})`); seen[h]++; } });
  return data.map(row => H.reduce((acc, h, i) => { acc[h] = row[i] || ''; return acc; }, {}));
}

app.post('/gestion-call/sync', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  try {
    await ensureGestionCallSchema();
    const { rows: cnt } = await pgPool.query("SELECT COUNT(*)::int AS n FROM gestion_call WHERE origen = 'form'");
    if (cnt[0].n > 0 && req.query.force !== '1') return res.json({ success: true, yaMigrado: true, existentes: cnt[0].n });
    const data = await leerCallSheet();
    const filas = data.filter(r =>
      (r['Marca temporal'] || '').toString().trim() !== '' || (r['ASESOR CONTACT'] || '').toString().trim() !== '');
    const client = await pgPool.connect();
    let insertados = 0;
    try {
      await client.query('BEGIN');
      const CHUNK = 500;
      for (let i = 0; i < filas.length; i += CHUNK) {
        const chunk = filas.slice(i, i + CHUNK);
        const params = [];
        const tuples = chunk.map((r, idx) => {
          const arr = mapCallRow(r, 'form');
          const base = idx * GC_COLS.length;
          params.push(...arr);
          return '(' + GC_COLS.map((_, j) => `$${base + j + 1}`).join(',') + ')';
        });
        await client.query(`INSERT INTO gestion_call (${GC_COLS.join(',')}) VALUES ${tuples.join(',')}`, params);
        insertados += chunk.length;
      }
      await client.query('COMMIT');
    } catch (e) { await client.query('ROLLBACK'); throw e; } finally { client.release(); }
    res.json({ success: true, leidas: data.length, insertados });
  } catch (e) { console.error('❌ POST /gestion-call/sync:', e); res.status(500).json({ success: false, message: e.message }); }
});

app.post('/gestion-call', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  const b = req.body || {};
  if (!b.asesor_contact || !b.dni_cliente || !b.estado_gestion) {
    return res.status(400).json({ success: false, message: 'Faltan campos obligatorios (asesor, dni, estado de gestión).' });
  }
  try {
    await ensureGestionCallSchema();
    const t = ahoraLima();
    const valorDe = {
      marca_temporal: t.ts, marca_temporal_raw: t.raw, origen: 'app',
      asesor_contact: b.asesor_contact, dni_cliente: b.dni_cliente, tipo_cliente: b.tipo_cliente,
      estado_gestion: b.estado_gestion, medio_primer_contacto: b.medio_primer_contacto,
      celular_gestionado: b.celular_gestionado, resultado_gestion: b.resultado_gestion,
      producto_interes: b.producto_interes, motivo_interes: b.motivo_interes, motivo_agendamiento: b.motivo_agendamiento,
      fecha_interes_agendamiento: b.fecha_interes_agendamiento, hora_interes_agendamiento: b.hora_interes_agendamiento,
      fecha_interes_derivacion: b.fecha_interes_derivacion, hora_interes_derivacion: b.hora_interes_derivacion,
      comentario_derivacion: b.comentario_derivacion, comentario_agendamiento: b.comentario_agendamiento,
      motivo_no_interes: b.motivo_no_interes, comentario_no_interes: b.comentario_no_interes,
      motivo_no_atendible: b.motivo_no_atendible, comentario_no_atendible: b.comentario_no_atendible,
      motivos_tercero_relacionado: b.motivos_tercero_relacionado, fecha_rellamada: b.fecha_rellamada,
      hora_rellamada: b.hora_rellamada, numero_titular_actual: b.numero_titular_actual,
      motivo_no_contacto: b.motivo_no_contacto, sede: b.sede, kommo: b.kommo,
      motivo_no_cierre: b.motivo_no_cierre, comentario_venta_no_concretada: b.comentario_venta_no_concretada,
    };
    const params = GC_COLS.map(c => { const v = valorDe[c]; return v === undefined || v === '' ? null : v; });
    const ph = GC_COLS.map((_, i) => `$${i + 1}`).join(',');
    const { rows } = await pgPool.query(`INSERT INTO gestion_call (${GC_COLS.join(',')}) VALUES (${ph}) RETURNING id`, params);
    res.json({ success: true, id: rows[0].id, marca_temporal: valorDe.marca_temporal_raw });
  } catch (e) { console.error('❌ POST /gestion-call:', e); res.status(500).json({ success: false, message: 'No se pudo guardar la gestión.' }); }
});

app.get('/gestion-call', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  try {
    await ensureGestionCallSchema();
    const cond = []; const params = [];
    if (req.query.desde) { params.push(`${req.query.desde} 00:00:00`); cond.push(`marca_temporal >= $${params.length}`); }
    if (req.query.hasta) { params.push(`${req.query.hasta} 23:59:59`); cond.push(`marca_temporal <= $${params.length}`); }
    const where = cond.length ? 'WHERE ' + cond.join(' AND ') : '';
    const { rows } = await pgPool.query(`SELECT * FROM gestion_call ${where} ORDER BY marca_temporal DESC NULLS LAST`, params);
    res.json(rows.map(gcRowToSheet));
  } catch (e) { console.error('❌ GET /gestion-call:', e); res.status(500).json({ success: false, message: 'No se pudieron obtener las gestiones.' }); }
});

// Mapa cabecera de hoja → columna BD (Call), para editar desde el grid.
const CALL_SHEET_TO_COL = {
  'ASESOR CONTACT': 'asesor_contact', 'DNI CLIENTE': 'dni_cliente', 'TIPO DE CLIENTE': 'tipo_cliente',
  'ESTADO DE GESTIÓN': 'estado_gestion', 'MEDIO DE PRIMER CONTACTO': 'medio_primer_contacto', 'CELULAR GESTIONADO': 'celular_gestionado',
  'RESULTADO DE GESTIÓN': 'resultado_gestion', 'PRODUCTO INTERÉS': 'producto_interes', 'MOTIVO INTERÉS': 'motivo_interes',
  'MOTIVO AGENDAMIENTO': 'motivo_agendamiento', 'FECHA DE INTERÉS AGENDAMIENTO': 'fecha_interes_agendamiento',
  'HORA APROXIMADA INTERÉS AGENDAMIENTO': 'hora_interes_agendamiento', 'FECHA DE INTERÉS DERIVACIÓN': 'fecha_interes_derivacion',
  'HORA APROXIMADA INTERÉS DERIVACIÓN': 'hora_interes_derivacion', 'COMENTARIO ADICIONAL DERIVACIÓN': 'comentario_derivacion',
  'COMENTARIO ADICIONAL AGENDAMIENTO': 'comentario_agendamiento', 'MOTIVO NO INTERÉS': 'motivo_no_interes',
  'COMENTARIO ADICIONAL NO INTERES': 'comentario_no_interes', 'MOTIVO NO ATENDIBLE': 'motivo_no_atendible',
  'COMENTARIO ADICIONAL NO ATENDIBLE': 'comentario_no_atendible', 'MOTIVOS TERCERO RELACIONADO': 'motivos_tercero_relacionado',
  'FECHA DE RE-LLAMADA': 'fecha_rellamada', 'HORA DE RELLAMADA': 'hora_rellamada', 'NÚMERO TITULAR ACTUAL': 'numero_titular_actual',
  'MOTIVO NO CONTACTO': 'motivo_no_contacto', 'SEDE': 'sede', 'KOMMO': 'kommo',
  'MOTIVO DE NO CIERRE': 'motivo_no_cierre', 'COMENTARIO VENTA NO CONCRETADA': 'comentario_venta_no_concretada',
};

// PUT /gestion-call/:id — edita una gestión.
app.put('/gestion-call/:id', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  try {
    await ensureGestionCallSchema();
    const { sets, params } = construirUpdate(req.body, CALL_SHEET_TO_COL, GC_COLS);
    if (!sets.length) return res.status(400).json({ success: false, message: 'Nada para actualizar.' });
    params.push(parseInt(req.params.id, 10));
    const { rowCount } = await pgPool.query(`UPDATE gestion_call SET ${sets.join(', ')} WHERE id = $${params.length}`, params);
    if (!rowCount) return res.status(404).json({ success: false, message: 'Gestión no encontrada.' });
    res.json({ success: true });
  } catch (e) { console.error('❌ PUT /gestion-call/:id', e); res.status(500).json({ success: false, message: 'No se pudo actualizar.' }); }
});

// DELETE /gestion-call/:id — elimina una gestión.
app.delete('/gestion-call/:id', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  try {
    await ensureGestionCallSchema();
    const { rowCount } = await pgPool.query('DELETE FROM gestion_call WHERE id = $1', [parseInt(req.params.id, 10)]);
    if (!rowCount) return res.status(404).json({ success: false, message: 'Gestión no encontrada.' });
    res.json({ success: true });
  } catch (e) { console.error('❌ DELETE /gestion-call/:id', e); res.status(500).json({ success: false, message: 'No se pudo eliminar.' }); }
});

// POST /gestion-call/match — { dnis: [...] } → { <dni>: {asesor, tipo_cliente, sede, kommo, celular} }
// Devuelve la ÚLTIMA gestión Call de cada DNI (para atribuir las ventas del Excel).
app.post('/gestion-call/match', async (req, res) => {
  if (!pgPool) return res.status(500).json({ success: false, message: 'Base de datos no configurada.' });
  try {
    await ensureGestionCallSchema();
    const dnis = Array.from(new Set((req.body?.dnis || []).map(d => String(d).replace(/\D/g, '').replace(/^0+/, '')).filter(Boolean)));
    if (!dnis.length) return res.json({});
    const { rows } = await pgPool.query(`
      SELECT DISTINCT ON (dnin) dnin, asesor_contact, tipo_cliente, sede, kommo, celular_gestionado
      FROM (
        SELECT regexp_replace(regexp_replace(dni_cliente, '\\D', '', 'g'), '^0+', '') AS dnin,
               asesor_contact, tipo_cliente, sede, kommo, celular_gestionado, marca_temporal
        FROM gestion_call
        WHERE regexp_replace(regexp_replace(dni_cliente, '\\D', '', 'g'), '^0+', '') = ANY($1)
      ) t
      ORDER BY dnin, marca_temporal DESC NULLS LAST
    `, [dnis]);
    const map = {};
    rows.forEach(r => { map[r.dnin] = { asesor: r.asesor_contact || '', tipo_cliente: r.tipo_cliente || '', sede: r.sede || '', kommo: r.kommo || '', celular: r.celular_gestionado || '' }; });
    res.json(map);
  } catch (e) { console.error('❌ POST /gestion-call/match', e); res.status(500).json({ success: false, message: 'No se pudo hacer el match.' }); }
});

// Iniciar servidor
app.listen(PORT, () => {
  console.log(`✅ API corriendo en http://localhost:${PORT}`);
  Promise.all([ensureVentasSchema(), ensureMargenSchema(), ensureUsuariosSchema(), ensurePermisosSchema(), ensureGestionRealzzaSchema(), ensureGestionCallSchema()])
    .then(async () => {
      if (!pgPool) return;
      console.log('🐘 Esquemas verificados (ventas, margen, usuarios, permisos, gestión realzza, gestión call).');
      await migrarUsuariosDesdeSheet();
    })
    .catch((e) => console.error('❌ No se pudo verificar el esquema:', e));
});

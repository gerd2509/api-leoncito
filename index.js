require('dotenv').config();
const express = require('express');
const { google } = require('googleapis');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

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
    const jsonData = data.map((row) =>
      headers.reduce((acc, header, i) => ({
        ...acc,
        [header]: row[i] || '',
      }), {})
    );

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

// Iniciar servidor
app.listen(PORT, () => {
  console.log(`✅ API corriendo en http://localhost:${PORT}`);
});

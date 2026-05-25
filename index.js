const express = require('express');
const { google } = require('googleapis');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());

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

// Iniciar servidor
app.listen(PORT, () => {
  console.log(`✅ API corriendo en http://localhost:${PORT}`);
});

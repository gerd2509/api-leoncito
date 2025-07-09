const express = require('express');
const { google } = require('googleapis');
const cors = require('cors');
const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
  scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
});

app.get('/data', async (req, res) => {
  try {
    const client = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: client });

    const spreadsheetId = '1j3b7k-vD9UzWLqz6JJksm5Vj3dWvtqL4SckMP21II94'; // Reemplaza por tu ID real
    const range = 'Respuestas de formulario 1!A1:Z'; // Ajusta si usas otro nombre de hoja

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      return res.status(404).send('No data found.');
    }

    // Separar encabezados y datos
    const [rawHeaders, ...data] = rows;

    // Renombrar encabezados duplicados
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

    // Mapear los datos con los encabezados únicos
    const jsonData = data.map(row =>
      headers.reduce((acc, header, i) => ({
        ...acc,
        [header]: row[i] || ''
      }), {})
    );

    res.json(jsonData);
  } catch (error) {
    console.error('Error al obtener datos de Google Sheets:', error);
    res.status(500).send('Error al obtener datos de Google Sheets');
  }
});

app.listen(PORT, () => {
  console.log(`API corriendo en puerto ${PORT}`);
});
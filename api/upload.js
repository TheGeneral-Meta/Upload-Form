const XLSX = require('xlsx');
const { google } = require('googleapis');

// Helper: parse private key dengan benar
function parsePrivateKey(key) {
  // Handle jika key sudah dalam format yang benar
  if (key.includes('-----BEGIN PRIVATE KEY-----')) {
    return key;
  }
  // Jika key dalam format string dengan \n
  return key.replace(/\\n/g, '\n');
}

module.exports = async (req, res) => {
  // CORS
  res.setHeader('Access-Control-Allow-Credentials', true);
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,PATCH,DELETE,POST,PUT');
  res.setHeader('Access-Control-Allow-Headers', 'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { credentials, fileContent, fileName, mode, spreadsheetId: existingSpreadsheetId } = req.body;
    
    if (!credentials || !fileContent) {
      return res.status(400).json({ error: 'Missing credentials or file' });
    }

    // Parse credentials
    let creds;
    try {
      creds = typeof credentials === 'string' ? JSON.parse(credentials) : credentials;
    } catch (e) {
      return res.status(400).json({ error: 'Invalid credentials JSON' });
    }

    // Validasi credentials
    if (!creds.client_email || !creds.private_key) {
      return res.status(400).json({ error: 'Invalid credentials: missing client_email or private_key' });
    }

    // Setup Google Auth
    const auth = new google.auth.JWT(
      creds.client_email,
      null,
      parsePrivateKey(creds.private_key),
      ['https://www.googleapis.com/auth/spreadsheets']
    );

    const sheets = google.sheets({ version: 'v4', auth });

    // Parse file Excel dari base64
    const buffer = Buffer.from(fileContent, 'base64');
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const sheetNames = workbook.SheetNames;

    let targetSpreadsheetId = existingSpreadsheetId;

    // Buat spreadsheet baru jika mode = new
    if (mode === 'new' || !targetSpreadsheetId) {
      const createResponse = await sheets.spreadsheets.create({
        requestBody: {
          properties: {
            title: `Rekap Data ${new Date().toLocaleDateString('id-ID')}`
          }
        }
      });
      targetSpreadsheetId = createResponse.data.spreadsheetId;
    }

    // Upload setiap sheet
    const uploadedSheets = [];
    for (const sheetName of sheetNames) {
      const worksheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
      
      if (!data || data.length === 0) continue;

      // Cek apakah sheet sudah ada
      let sheetExists = false;
      try {
        const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: targetSpreadsheetId });
        sheetExists = spreadsheet.data.sheets.some(s => s.properties.title === sheetName);
      } catch (e) {
        // Spreadsheet mungkin baru dibuat
      }

      if (!sheetExists) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: targetSpreadsheetId,
          requestBody: {
            requests: [{ addSheet: { properties: { title: sheetName } } }]
          }
        });
      }

      // Clear data existing (opsional)
      try {
        await sheets.spreadsheets.values.clear({
          spreadsheetId: targetSpreadsheetId,
          range: `${sheetName}!A1:ZZZ`
        });
      } catch (e) {}

      // Tulis data baru
      if (data.length > 0) {
        await sheets.spreadsheets.values.update({
          spreadsheetId: targetSpreadsheetId,
          range: `${sheetName}!A1`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: data }
        });
      }
      
      uploadedSheets.push(sheetName);
    }

    // Share spreadsheet ke email tertentu (opsional)
    if (creds.shareWithEmail) {
      try {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: targetSpreadsheetId,
          requestBody: {
            requests: [{
              addDomain: {
                domain: creds.shareWithEmail.split('@')[1]
              }
            }]
          }
        });
      } catch (e) {}
    }

    return res.json({
      success: true,
      message: `✅ Berhasil! ${uploadedSheets.length} sheet terupload.`,
      spreadsheetId: targetSpreadsheetId,
      url: `https://docs.google.com/spreadsheets/d/${targetSpreadsheetId}`,
      sheets: uploadedSheets
    });

  } catch (error) {
    console.error('Error:', error);
    return res.status(500).json({ 
      success: false, 
      error: error.message || 'Terjadi kesalahan',
      details: error.toString()
    });
  }
};

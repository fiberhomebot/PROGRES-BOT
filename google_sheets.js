const { google } = require('googleapis');

class GoogleSheets {
  constructor({ spreadsheetId, sheetName }) {
    this.spreadsheetId = spreadsheetId;
    this.sheetName = sheetName || 'Sheet1';
    this.sheets = null;
    this.authClient = null;
    // start async init and store promise so callers can await readiness
    this.initPromise = this.init();
  }

  async init() {
    const credsEnv = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
    if (!credsEnv) {
      console.error('Missing GOOGLE_SERVICE_ACCOUNT_JSON env var');
      return;
    }
    let creds;
    try {
      creds = JSON.parse(credsEnv);
    } catch (err) {
      console.error('Failed parsing GOOGLE_SERVICE_ACCOUNT_JSON');
      throw err;
    }

    const scopes = ['https://www.googleapis.com/auth/spreadsheets'];
    const jwt = new google.auth.JWT(creds.client_email, null, creds.private_key, scopes);
    this.authClient = jwt;
    this.sheets = google.sheets({ version: 'v4', auth: this.authClient });
  }

  async getHeaders() {
    await this.initPromise;
    const range = `${this.sheetName}!1:1`;
    const res = await this.sheets.spreadsheets.values.get({ spreadsheetId: this.spreadsheetId, range });
    const values = res.data.values || [];
    return values[0] || [];
  }

  async appendRow(obj) {
    await this.initPromise;
    const headers = await this.getHeaders();
    const row = new Array(headers.length).fill('');
    const newHeaders = [];

    for (const key of Object.keys(obj)) {
      const k = key.toUpperCase();
      const idx = headers.findIndex(h => (h || '').toUpperCase() === k);
      if (idx !== -1) {
        row[idx] = obj[key];
      } else {
        newHeaders.push({ key: k, value: obj[key] });
      }
    }

    // if there are new headers, append them to header row
    if (newHeaders.length > 0) {
      const newHeaderNames = newHeaders.map(n => n.key);
      const combinedHeaders = headers.concat(newHeaderNames);
      // update the entire header row so new columns are in correct order
      await this.sheets.spreadsheets.values.update({
        spreadsheetId: this.spreadsheetId,
        range: `${this.sheetName}!1:1`,
        valueInputOption: 'RAW',
        requestBody: { values: [combinedHeaders] }
      });
      // extend the row array to match combined headers
      for (const n of newHeaders) row.push(n.value);
    }

    // append row
    await this.sheets.spreadsheets.values.append({
      spreadsheetId: this.spreadsheetId,
      range: this.sheetName,
      valueInputOption: 'RAW',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [row] }
    });
  }

  async getAllRows() {
    await this.initPromise;
    const range = this.sheetName;
    const res = await this.sheets.spreadsheets.values.get({ spreadsheetId: this.spreadsheetId, range });
    return res.data.values || [];
  }

  async ensureHeader(headerName) {
    await this.initPromise;
    const headers = await this.getHeaders();
    const up = (h) => (h||'').toString().toUpperCase();
    const idx = headers.findIndex(h => up(h) === headerName.toUpperCase());
    if (idx !== -1) return idx;
    // append header
    const newHeaders = headers.concat([headerName]);
    await this.sheets.spreadsheets.values.update({
      spreadsheetId: this.spreadsheetId,
      range: `${this.sheetName}!1:1`,
      valueInputOption: 'RAW',
      requestBody: { values: [newHeaders] }
    });
    return newHeaders.findIndex(h => up(h) === headerName.toUpperCase());
  }

  async updateCell(rowIndex, colIndex, value) {
    await this.initPromise;
    const a1 = `${this.sheetName}!${this._colToLetter(colIndex)}${rowIndex}`;
    await this.sheets.spreadsheets.values.update({
      spreadsheetId: this.spreadsheetId,
      range: a1,
      valueInputOption: 'RAW',
      requestBody: { values: [[value]] }
    });
  }

  _colToLetter(col) {
    let letter = '';
    while (col > 0) {
      const mod = (col - 1) % 26;
      letter = String.fromCharCode(65 + mod) + letter;
      col = Math.floor((col - mod) / 26);
    }
    return letter;
  }
}

module.exports = { GoogleSheets };

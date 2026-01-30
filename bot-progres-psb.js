require('dotenv').config();
// Normalize environment variable names for compatibility with older setups
process.env.TELEGRAM_TOKEN = process.env.TELEGRAM_TOKEN || process.env.BOT_TOKEN;
process.env.SHEET_ID = process.env.SHEET_ID || process.env.SPREADSHEET_ID;
process.env.GOOGLE_SERVICE_ACCOUNT_KEY = process.env.GOOGLE_SERVICE_ACCOUNT_KEY || process.env.GOOGLE_SERVICE_ACCOUNT_JSON || process.env.GOOGLE_SERVICE_ACCOUNT;
const TelegramBot = require('node-telegram-bot-api');
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

// Inline parser function untuk PROGRES PSB
function parseProgres(text, userRow, username) {
  const lines = text.split('\n').map(l=>l.trim()).filter(l=>l);
  const upper = text.toUpperCase();
  
  let channel='', dateCreated='', workorder='', scOrderNo='', ao='', serviceNo='', 
      address='', customerName='', workzone='', contactPhone='', bookingDate='', 
      odp='', paket='', mitra='', symptom='', memo='', tikor='', ncli='', teknisi='';
  
  teknisi = (userRow && userRow[1]) ? (userRow[1] || username).replace('@', '') : (username || '');

  function findValue(patterns) {
    for (const pattern of patterns) {
      const matches = text.match(pattern);
      if (matches) {
        if (pattern.global) {
          return matches[matches.length - 1];
        } else if (matches[1]) {
          return matches[1].trim();
        }
      }
    }
    return '';
  }

  function detectChannel(text) {
    const upperText = text.toUpperCase();
    if (upperText.includes('CHANNEL : DIGIPOS') || upperText.includes('CHANNEL: DIGIPOS')) {
      return 'DIGIPOS';
    }
    if (upperText.includes('CHANNEL : BS') || upperText.includes('CHANNEL: BS') ||
        upperText.includes('CHANNEL : GS') || upperText.includes('CHANNEL: GS') ||
        upperText.includes('CHANNEL : ES') || upperText.includes('CHANNEL: ES')) {
      const match = text.match(/CHANNEL\s*:\s*([A-Z]+)/i);
      return match && match[1] ? match[1] : '';
    }
    return '';
  }

  channel = detectChannel(text);

  // Generic pattern parsing untuk kedua channel
  channel = findValue([/CHANNEL\s*:\s*([A-Z]+)/i]) || channel;
  dateCreated = findValue([/DATE\s*CREATED\s*:\s*(.+?)(?=\n|$)/i]);
  workorder = findValue([/WORKORDER\s*:\s*([A-Za-z0-9]+)/i]);
  scOrderNo = findValue([/SC\s*ORDER\s*NO\s*:\s*([A-Za-z0-9\-]+)/i]);
  ao = findValue([/AO\s*:\s*([A-Za-z0-9]+)/i]);
  serviceNo = findValue([/SERVICE\s*NO\s*:\s*(\d+)/i]);
  address = findValue([/ADDRESS\s*:\s*(.+?)(?=\n[A-Z]+\s*:|$)/i]);
  customerName = findValue([/CUSTOMER\s*NAME\s*:\s*(.+?)(?=\n|$)/i]);
  workzone = findValue([/WORKZONE\s*:\s*([A-Z0-9]+)/i]);
  contactPhone = findValue([/CONTACT\s*PHONE\s*:\s*([0-9\+\-\s]+)/i]);
  bookingDate = findValue([/BOOKING\s*DATE\s*:\s*(.+?)(?=\n|$)/i]);
  odp = findValue([/ODP\s*:\s*(.+?)(?=\n|$)/i]);
  paket = findValue([/PAKET\s*:\s*(.+?)(?=\n|$)/i]) || findValue([/PACKAGE\s*:\s*(.+?)(?=\n|$)/i]);
  mitra = findValue([/MITRA\s*:\s*(.+?)(?=\n|$)/i]);
  symptom = findValue([/SYMPTOM\s*:\s*(.+?)(?=\n|$)/i]);
  memo = findValue([/MEMO\s*:\s*(.+?)(?=\n|$)/i]);
  tikor = findValue([/TIKOR\s*:\s*(.+?)(?=\n|$)/i]);
  ncli = findValue([/NCLI\s*:\s*([0-9]+)/i]);

  return { 
    channel, dateCreated, workorder, scOrderNo, ao, serviceNo, address, 
    customerName, workzone, contactPhone, bookingDate, odp, paket, mitra, 
    symptom, memo, tikor, ncli, teknisi 
  };
}

// === Konfigurasi dari environment variables ===
const TOKEN = process.env.TELEGRAM_TOKEN;
const SHEET_ID = process.env.SHEET_ID;
let GOOGLE_SERVICE_ACCOUNT_KEY = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;

// Validasi environment variables
if (!TOKEN) {
  console.error('ERROR: TELEGRAM_TOKEN environment variable is not set!');
  process.exit(1);
}
if (!SHEET_ID) {
  console.error('ERROR: SHEET_ID environment variable is not set!');
  process.exit(1);
}

// Try to read from file first if env var is not set or looks like a path
if (!GOOGLE_SERVICE_ACCOUNT_KEY || /\.json$/i.test(GOOGLE_SERVICE_ACCOUNT_KEY)) {
  const credentialPath = GOOGLE_SERVICE_ACCOUNT_KEY || './service-account.json';
  try {
    if (fs.existsSync(credentialPath)) {
      GOOGLE_SERVICE_ACCOUNT_KEY = fs.readFileSync(credentialPath, 'utf8');
      console.log(`Loaded service account from file: ${credentialPath}`);
    }
  } catch (e) {
    console.log(`Could not read credentials from ${credentialPath}, will try env var`);
  }
}

if (!GOOGLE_SERVICE_ACCOUNT_KEY) {
  console.error('ERROR: GOOGLE_SERVICE_ACCOUNT_KEY environment variable is not set and no service-account.json file found!');
  process.exit(1);
}

const PROGRES_SHEET = 'PROGRES PSB';
const USER_SHEET = 'USER';

// === Setup Google Sheets API ===
let serviceAccount;
try {
  let keyData = (GOOGLE_SERVICE_ACCOUNT_KEY || '').toString();

  // Trim and remove BOM if present
  keyData = keyData.trim();
  if (keyData.charCodeAt(0) === 0xFEFF) {
    keyData = keyData.slice(1);
  }

  // If the value looks like a path to a JSON file, try reading it
  try {
    const looksLikePath = /(^\.|^\\|^\/|\\.json$|\.json$)/i.test(keyData);
    if (looksLikePath) {
      try {
        const possiblePath = keyData.replace(/^file:\/\//i, '');
        if (fs.existsSync(possiblePath)) {
          keyData = fs.readFileSync(possiblePath, 'utf8');
        }
      } catch (e) {
        // ignore and continue trying other strategies
      }
    }
  } catch (e) {
    // ignore
  }

  // If it's not a direct JSON string, try common transformations
  if (!keyData.startsWith('{')) {
    // Try base64 decode
    try {
      const decoded = Buffer.from(keyData, 'base64').toString('utf-8');
      if (decoded.trim().startsWith('{')) {
        keyData = decoded;
      }
    } catch (e) {
      // not base64 or decode failed
    }
  }

  // If still not JSON, maybe it contains escaped newlines; unescape and trim surrounding quotes
  if (!keyData.trim().startsWith('{')) {
    let unescaped = keyData.replace(/\\n/g, '\n').replace(/\\\"/g, '\"');
    if ((unescaped.startsWith('"') && unescaped.endsWith('"')) || (unescaped.startsWith("'") && unescaped.endsWith("'"))) {
      unescaped = unescaped.slice(1, -1);
    }
    keyData = unescaped;
  }

  // Final parse attempt
  serviceAccount = JSON.parse(keyData);

  // Log safely: show a small prefix (masked) to help debugging without printing secret
  const preview = JSON.stringify(Object.keys(serviceAccount || {})).slice(0, 200);
  console.log('Google Service Account parsed successfully; keys:', preview);
} catch (e) {
  const sample = (GOOGLE_SERVICE_ACCOUNT_KEY || '').toString().slice(0, 120).replace(/\n/g, '\\n');
  console.error('ERROR parsing GOOGLE_SERVICE_ACCOUNT_KEY:', e.message);
  console.error('Sample of provided value (truncated, escaped):', sample);
  console.error('Hint: set GOOGLE_SERVICE_ACCOUNT_KEY to the raw JSON content, or a base64-encoded JSON string, without extra surrounding quotes.');
  process.exit(1);
}

const auth = new google.auth.GoogleAuth({
  credentials: serviceAccount,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });

// === Setup Telegram Bot dengan webhook untuk Railway ===
let bot;
const PORT = process.env.PORT || 3001;
const RAILWAY_STATIC_URL = process.env.RAILWAY_STATIC_URL;
const USE_WEBHOOK = process.env.USE_WEBHOOK === 'true' || !!RAILWAY_STATIC_URL;

if (USE_WEBHOOK && RAILWAY_STATIC_URL) {
  const express = require('express');
  const app = express();
  app.use(express.json());
  
  bot = new TelegramBot(TOKEN);
  const webhookUrl = `https://${RAILWAY_STATIC_URL}/progres${TOKEN}`;
  
  bot.setWebHook(webhookUrl).then(() => {
    console.log(`Webhook set to: ${webhookUrl}`);
  }).catch(err => {
    console.error('Failed to set webhook:', err);
  });
  
  app.post(`/progres${TOKEN}`, (req, res) => {
    bot.processUpdate(req.body);
    res.sendStatus(200);
  });
  
  app.get('/', (req, res) => {
    res.send('Bot Progres PSB is running!');
  });
  
  app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
  });
} else {
  bot = new TelegramBot(TOKEN, { polling: true });
  console.log('Bot Progres PSB running in polling mode');
}

// === Helper: Ambil data dari sheet ===
async function getSheetData(sheetName) {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: sheetName,
    });
    return res.data.values || [];
  } catch (error) {
    console.error(`Error getting sheet data from ${sheetName}:`, error.message);
    throw error;
  }
}

// === Helper: Tambah data ke sheet ===
async function appendSheetData(sheetName, values) {
  try {
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: sheetName,
      valueInputOption: 'USER_ENTERED',
      resource: { values: [values] },
    });
  } catch (error) {
    console.error(`Error appending data to ${sheetName}:`, error.message);
    throw error;
  }
}

// === Helper: Update range sheet data ===
async function updateSheetData(sheetName, range, values) {
  try {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${sheetName}!${range}`,
      valueInputOption: 'USER_ENTERED',
      resource: { values },
    });
  } catch (error) {
    console.error(`Error updating sheet data:`, error.message);
    throw error;
  }
}

// === Helper: Kirim pesan Telegram dengan retry logic ===
async function sendTelegram(chatId, text, options = {}) {
  const maxLength = 4000;
  const maxRetries = 3;
  
  async function sendWithRetry(message, retries = 0) {
    try {
      return await bot.sendMessage(chatId, message, { parse_mode: 'HTML', ...options });
    } catch (error) {
      if (retries < maxRetries) {
        console.log(`Retry ${retries + 1} sending message to ${chatId}`);
        await new Promise(resolve => setTimeout(resolve, 1000 * (retries + 1)));
        return sendWithRetry(message, retries + 1);
      }
      throw error;
    }
  }
  
  if (text.length <= maxLength) {
    return sendWithRetry(text);
  } else {
    const lines = text.split('\n');
    let chunk = '';
    let promises = [];
    for (let i = 0; i < lines.length; i++) {
      if ((chunk + lines[i] + '\n').length > maxLength) {
        promises.push(sendWithRetry(chunk));
        chunk = '';
      }
      chunk += lines[i] + '\n';
    }
    if (chunk.trim()) promises.push(sendWithRetry(chunk));
    return Promise.all(promises);
  }
}

// === Helper: Cek user aktif ===
async function getUserData(username) {
  try {
    const data = await getSheetData(USER_SHEET);
    for (let i = 1; i < data.length; i++) {
      const userSheetUsername = (data[i][1] || '').replace('@', '').toLowerCase();
      const inputUsername = (username || '').replace('@', '').toLowerCase();
      const userStatus = (data[i][3] || '').toUpperCase();
      if (userSheetUsername === inputUsername && userStatus === 'AKTIF') {
        return data[i];
      }
    }
    return null;
  } catch (error) {
    console.error('Error getting user data:', error);
    return null;
  }
}

// === Helper: Cek admin ===
async function isAdmin(username) {
  const user = await getUserData(username);
  return user && (user[2] || '').toUpperCase() === 'ADMIN';
}

// === Helper: Get today's date string ===
function getTodayDateString() {
  const today = new Date();
  return today.toLocaleDateString('id-ID', {
    weekday: 'long',
    day: 'numeric',
    month: 'long',
    year: 'numeric',
    timeZone: 'Asia/Jakarta'
  });
}

// === Handler pesan masuk ===
bot.on('message', async (msg) => {
  const chatId = msg.chat.id;
  const messageId = msg.message_id;
  const text = (msg.text || '').trim();
  const username = msg.from.username || '';
  const chatType = msg.chat.type;
  
  console.log(`Message received - Chat: ${chatId}, User: @${username}, Type: ${chatType}, Text: ${text.substring(0, 50)}`);
  
  try {
    // Hanya proses /UPDATE di group, command lain diabaikan
    if ((chatType === 'group' || chatType === 'supergroup') && !/^\/UPDATE\b/i.test(text)) {
      return;
    }
    
    // === /UPDATE: Input data progres PSB ===
    if (/^\/UPDATE\b/i.test(text)) {
      const user = await getUserData(username);
      if (!user) {
        return sendTelegram(chatId, '❌ Anda tidak terdaftar sebagai user aktif.', { reply_to_message_id: messageId });
      }
      
      const inputText = text.replace(/^\/UPDATE\s*/i, '').trim();
      if (!inputText) {
        return sendTelegram(chatId, 'Silakan kirim data progres setelah /UPDATE.', { reply_to_message_id: messageId });
      }
      
      // Parse data
      const parsed = parseProgres(inputText, user, username);
      
      // Validasi minimal field penting
      let missing = [];
      if (!parsed.channel) missing.push('CHANNEL');
      if (!parsed.scOrderNo) missing.push('SC ORDER NO');
      if (!parsed.serviceNo) missing.push('SERVICE NO');
      if (!parsed.customerName) missing.push('CUSTOMER NAME');
      if (!parsed.workzone) missing.push('WORKZONE');
      
      if (missing.length > 0) {
        return sendTelegram(chatId, `❌ Data tidak lengkap. Field wajib: ${missing.join(', ')}`, { reply_to_message_id: messageId });
      }
      
      // Susun data sesuai urutan kolom sheet PROGRES PSB
      const row = [
        parsed.dateCreated || getTodayDateString(),  // DATE CREATED
        parsed.channel,                               // CHANNEL
        parsed.workorder,                             // WORKORDER
        parsed.scOrderNo,                             // SC ORDER NO
        parsed.serviceNo,                             // SERVICE NO
        parsed.customerName,                          // CUSTOMER NAME
        parsed.workzone,                              // WORKZONE
        parsed.contactPhone,                          // CONTACT PHONE
        parsed.odp,                                   // ODP
        parsed.paket,                                 // PAKET
        parsed.memo,                                  // MEMO
        parsed.symptom,                               // SYMPTOM
        parsed.tikor,                                 // TIKOR
        parsed.address,                               // ADDRESS
        parsed.bookingDate,                           // BOOKING DATE
        parsed.mitra,                                 // MITRA
        parsed.ncli,                                  // NCLI
        parsed.ao,                                    // AO
        parsed.teknisi,                               // TEKNISI
        chatId,                                       // CHAT_ID (for tracking)
        messageId                                     // MESSAGE_ID (for tracking)
      ];
      
      await appendSheetData(PROGRES_SHEET, row);
      
      let confirmMsg = '✅ Data Progres PSB berhasil disimpan! 🚀\n\n';
      confirmMsg += '<b>DETAIL YANG DICATAT:</b>\n';
      confirmMsg += `📱 Channel: ${parsed.channel}\n`;
      confirmMsg += `🔢 SC Order: ${parsed.scOrderNo}\n`;
      confirmMsg += `👤 Customer: ${parsed.customerName}\n`;
      confirmMsg += `📍 Workzone: ${parsed.workzone}\n`;
      confirmMsg += `💬 Symptom: ${parsed.symptom || '-'}\n`;
      
      return sendTelegram(chatId, confirmMsg, { reply_to_message_id: messageId });
    }
    
    // === /progres: Laporan per teknisi dengan breakdown symptom ===
    else if (/^\/progres\b/i.test(text)) {
      const data = await getSheetData(PROGRES_SHEET);
      let teknisiMap = {}; // {teknisi: {symptom: count}}
      
      // Parse data dari sheet
      for (let i = 1; i < data.length; i++) {
        const teknisi = (data[i][17] || '-').trim(); // TEKNISI kolom 17
        const symptom = (data[i][15] || '-').trim(); // SYMPTOM kolom 15
        
        if (!teknisiMap[teknisi]) {
          teknisiMap[teknisi] = { total: 0 };
        }
        teknisiMap[teknisi].total++;
        teknisiMap[teknisi][symptom] = (teknisiMap[teknisi][symptom] || 0) + 1;
      }
      
      let msg = '📊 <b>LAPORAN PROGRES TEKNISI</b>\n\n';
      const sortedTeknisi = Object.entries(teknisiMap).sort((a, b) => b[1].total - a[1].total);
      
      sortedTeknisi.forEach((entry, idx) => {
        const [teknisi, counts] = entry;
        const total = counts.total;
        msg += `${idx + 1}. ${teknisi} : ${total} WO\n`;
        
        Object.entries(counts).forEach(([key, val]) => {
          if (key !== 'total') {
            msg += `   ${key} : ${val}\n`;
          }
        });
      });
      
      return sendTelegram(chatId, msg, { reply_to_message_id: messageId });
    }
    
    // === /cek: Total per workzone dengan breakdown symptom ===
    else if (/^\/cek\b/i.test(text)) {
      const data = await getSheetData(PROGRES_SHEET);
      let workzoneMap = {}; // {workzone: {symptom: count}}
      
      // Parse data dari sheet
      for (let i = 1; i < data.length; i++) {
        const workzone = (data[i][7] || '-').trim(); // WORKZONE kolom 7
        const symptom = (data[i][15] || '-').trim(); // SYMPTOM kolom 15
        
        if (!workzoneMap[workzone]) {
          workzoneMap[workzone] = { total: 0 };
        }
        workzoneMap[workzone].total++;
        workzoneMap[workzone][symptom] = (workzoneMap[workzone][symptom] || 0) + 1;
      }
      
      let msg = '📍 <b>REKAP PER WORKZONE</b>\n\n';
      const sortedWorkzone = Object.entries(workzoneMap).sort((a, b) => b[1].total - a[1].total);
      
      sortedWorkzone.forEach(([workzone, counts]) => {
        msg += `<b>${workzone} : ${counts.total}</b>\n`;
        
        Object.entries(counts)
          .sort((a, b) => b[1] - a[1])
          .forEach(([key, val]) => {
            if (key !== 'total') {
              msg += `  ${key} : ${val}\n`;
            }
          });
      });
      
      return sendTelegram(chatId, msg, { reply_to_message_id: messageId });
    }
    
    // === /[AO_value]: Cari berdasarkan AO (DIGIPOS) ===
    else if (/^\/[A-Za-z0-9]+$/i.test(text) && !text.match(/^\/cek|^\/progres|^\/help|^\/start|^\/UPDATE/i)) {
      const aoValue = text.substring(1).toUpperCase();
      const data = await getSheetData(PROGRES_SHEET);
      const results = [];
      
      // Search di kolom AO (index 3)
      for (let i = 1; i < data.length; i++) {
        const ao = (data[i][3] || '').toUpperCase();
        if (ao === aoValue) {
          results.push(data[i]);
        }
      }
      
      if (results.length === 0) {
        // Coba cari di SC ORDER NO (index 4) jika tidak ketemu di AO
        for (let i = 1; i < data.length; i++) {
          const scOrder = (data[i][4] || '').toUpperCase();
          if (scOrder === aoValue) {
            results.push(data[i]);
          }
        }
      }
      
      let msg = '';
      if (results.length === 0) {
        msg = `❌ AO/SC ORDER <b>${aoValue}</b> tidak ditemukan`;
      } else {
        results.forEach(order => {
          const ao = (order[3] || order[4] || '').trim(); // AO or SC ORDER NO
          const symptom = (order[15] || '-').trim(); // SYMPTOM
          msg += `<b>${ao}</b> : ${symptom}\n`;
        });
      }
      
      return sendTelegram(chatId, msg, { reply_to_message_id: messageId });
    }
    
    // === /help: Command list ===
    else if (/^\/help\b/i.test(text) || /^\/start\b/i.test(text)) {
      let helpMsg = '🤖 <b>Bot Progres PSB</b>\n\n';
      
      helpMsg += '<b>📝 COMMANDS UTAMA:</b>\n';
      helpMsg += '/UPDATE - Input data progres di group\n';
      helpMsg += '/progres - Laporan per teknisi + symptom breakdown\n';
      helpMsg += '/cek - Total per workzone + symptom breakdown\n';
      helpMsg += '/[AO] - Cek status order DIGIPOS\n';
      helpMsg += '/[SC_ORDER_NO] - Cek status order BS/GS/ES\n\n';
      
      helpMsg += '<b>📊 CONTOH PENGGUNAAN:</b>\n';
      helpMsg += '/progres\n';
      helpMsg += '/cek\n';
      helpMsg += '/AOi4260120124037284d1ca10\n';
      helpMsg += '/1002230081\n';
      
      return sendTelegram(chatId, helpMsg, { reply_to_message_id: messageId });
    }
    
    else if (text.startsWith('/')) {
      return sendTelegram(chatId, '❓ Command tidak dikenali. Ketik /help untuk melihat daftar command.', { reply_to_message_id: messageId });
    }
    
  } catch (err) {
    console.error('Error processing message:', err);
    return sendTelegram(chatId, '❌ Terjadi kesalahan sistem. Silakan coba lagi nanti.', { reply_to_message_id: messageId });
  }
});

// Error handling
process.on('uncaughtException', (err) => {
  console.error('Uncaught Exception:', err);
});

process.on('unhandledRejection', (reason, promise) => {
  console.error('Unhandled Rejection at:', promise, 'reason:', reason);
});

console.log('Bot Telegram Progres PSB started successfully!');
console.log('Mode:', USE_WEBHOOK ? 'Webhook' : 'Polling');
if (USE_WEBHOOK) {
  console.log('Listening on port:', PORT);
}

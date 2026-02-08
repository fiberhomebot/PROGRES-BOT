require('dotenv').config();
const TelegramBot = require('node-telegram-bot-api');
const { google } = require('googleapis');

// === ENV VARIABLES ===
const TOKEN = process.env.TELEGRAM_TOKEN;
const SHEET_ID = process.env.SHEET_ID;
let GOOGLE_SERVICE_ACCOUNT_JSON = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;

// Validasi
if (!TOKEN) {
  console.error('❌ TELEGRAM_TOKEN not set');
  process.exit(1);
}
if (!SHEET_ID) {
  console.error('❌ SHEET_ID not set');
  process.exit(1);
}
if (!GOOGLE_SERVICE_ACCOUNT_JSON) {
  console.error('❌ GOOGLE_SERVICE_ACCOUNT_JSON not set');
  process.exit(1);
}

// === PARSE GOOGLE SERVICE ACCOUNT ===
let serviceAccount;
try {
  let keyData = GOOGLE_SERVICE_ACCOUNT_JSON.trim();
  
  // Jika base64, decode dulu
  if (!keyData.startsWith('{')) {
    try {
      keyData = Buffer.from(keyData, 'base64').toString('utf-8');
    } catch (e) {
      // bukan base64
    }
  }
  
  serviceAccount = JSON.parse(keyData);
  console.log('✅ Google Service Account parsed');
} catch (e) {
  console.error('❌ Failed to parse JSON:', e.message);
  process.exit(1);
}

const auth = new google.auth.GoogleAuth({
  credentials: serviceAccount,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });

const PROGRES_SHEET = 'PROGRES PSB';
const MASTER_SHEET = 'MASTER';  // Changed from USER_SHEET

// === HELPER: Get sheet data ===
async function getSheetData(sheetName) {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: sheetName,
    });
    return res.data.values || [];
  } catch (error) {
    console.error(`Error reading ${sheetName}:`, error.message);
    throw error;
  }
}

// === HELPER: Append to sheet ===
async function appendSheetData(sheetName, values) {
  try {
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: sheetName,
      valueInputOption: 'USER_ENTERED',
      resource: { values: [values] },
    });
  } catch (error) {
    console.error(`Error writing to ${sheetName}:`, error.message);
    throw error;
  }
}

// === HELPER: Send Telegram ===
async function sendTelegram(chatId, text, options = {}) {
  const maxLength = 4000;
  try {
    if (text.length <= maxLength) {
      return await bot.sendMessage(chatId, text, { parse_mode: 'HTML', ...options });
    } else {
      const lines = text.split('\n');
      let chunk = '';
      for (let i = 0; i < lines.length; i++) {
        if ((chunk + lines[i] + '\n').length > maxLength) {
          await bot.sendMessage(chatId, chunk, { parse_mode: 'HTML', ...options });
          chunk = '';
        }
        chunk += lines[i] + '\n';
      }
      if (chunk.trim()) {
        await bot.sendMessage(chatId, chunk, { parse_mode: 'HTML', ...options });
      }
    }
  } catch (error) {
    console.error('Error sending message:', error.message);
  }
}

// === HELPER: Get user role ===
async function getUserRole(username) {
  try {
    const data = await getSheetData(MASTER_SHEET);
    for (let i = 1; i < data.length; i++) {
      const sheetUser = (data[i][8] || '').replace('@', '').toLowerCase().trim();
      const inputUser = (username || '').replace('@', '').toLowerCase().trim();
      const status = (data[i][10] || '').toUpperCase().trim();
      const role = (data[i][9] || '').toUpperCase().trim();
      
      if (sheetUser === inputUser && status === 'AKTIF') {
        return role;
      }
    }
    return null;
  } catch (error) {
    console.error('Error getting user role:', error.message);
    return null;
  }
}

// === HELPER: Get user data ===
async function getUserData(username) {
  try {
    const data = await getSheetData(MASTER_SHEET);
    for (let i = 1; i < data.length; i++) {
      const sheetUser = (data[i][8] || '').replace('@', '').toLowerCase().trim();
      const inputUser = (username || '').replace('@', '').toLowerCase().trim();
      const status = (data[i][10] || '').toUpperCase().trim();
      
      if (sheetUser === inputUser && status === 'AKTIF') {
        return data[i];
      }
    }
    return null;
  } catch (error) {
    console.error('Error getting user:', error.message);
    return null;
  }
}

// === HELPER: Check authorization ===
async function checkAuthorization(username, requiredRoles = []) {
  try {
    const userRole = await getUserRole(username);
    if (!userRole) {
      return { authorized: false, role: null, message: '❌ Anda tidak terdaftar di sistem.' };
    }
    
    if (requiredRoles.length > 0 && !requiredRoles.includes(userRole)) {
      return { authorized: false, role: userRole, message: `❌ Akses ditolak. Role ${userRole} tidak memiliki izin untuk command ini.` };
    }
    
    return { authorized: true, role: userRole };
  } catch (error) {
    return { authorized: false, role: null, message: '❌ Terjadi kesalahan saat verifikasi.' };
  }
}

// === HELPER: Parse progres data ===
function parseProgres(text, userRow, username) {
  let data = {
    channel: '',
    scOrderNo: '',
    serviceNo: '',
    customerName: '',
    workzone: '',
    contactPhone: '',
    odp: '',
    memo: '',
    symptom: '',
    ao: '',
    workorder: '',
    tikor: '',
    snOnt: '',
    nikOnt: '',
    stbId: '',
    nikStb: '',
    dateCreated: new Date().toLocaleDateString('id-ID', {
      weekday: 'long',
      day: 'numeric',
      month: 'long',
      year: 'numeric',
      timeZone: 'Asia/Jakarta',
    }),
    teknisi: (username || '').replace('@', ''),
  };

  const patterns = {
    channel: /CHANNEL\s*:\s*([A-Za-z0-9]+)/i,
    scOrderNo: /SC\s*ORDER\s*NO\s*:\s*(.+?)(?=\n|$)/i,
    serviceNo: /SERVICE\s*NO\s*:\s*([0-9]+)/i,
    customerName: /CUSTOMER\s*NAME\s*:\s*(.+?)(?=\n|$)/i,
    workzone: /WORKZONE\s*:\s*([A-Za-z0-9]+)/i,
    contactPhone: /CONTACT\s*PHONE\s*:\s*([0-9\+\-\s]+)/i,
    odp: /ODP\s*:\s*(.+?)(?=\n|$)/i,
    memo: /MEMO\s*:\s*(.+?)(?=\n|$)/i,
    symptom: /SYMPTOM\s*:\s*(.+?)(?=\n|$)/i,
    ao: /AO\s*:\s*(.+?)(?=\n|$)/i,
    workorder: /WORKORDER\s*:\s*([A-Za-z0-9]+)/i,
    tikor: /TIKOR\s*:\s*(.+?)(?=\n|$)/i,
    snOnt: /SN\s*ONT\s*:\s*(.+?)(?=\n|$)/i,
    nikOnt: /NIK\s*ONT\s*:\s*([0-9]+)/i,
    stbId: /STB\s*ID\s*:\s*(.+?)(?=\n|$)/i,
    nikStb: /NIK\s*STB\s*:\s*([0-9]+)/i,
  };

  for (const [key, pattern] of Object.entries(patterns)) {
    const match = text.match(pattern);
    if (match && match[1]) {
      data[key] = match[1].trim();
    }
  }

  return data;
}

// === HELPER: Parse aktivasi data ===
function parseAktivasi(text, username) {
  let data = {
    channel: '',
    dateCreated: '',
    scOrderNo: '',
    workorder: '',
    ao: '',
    ncli: '',
    serviceNo: '',
    address: '',
    customerName: '',
    workzone: '',
    contactPhone: '',
    bookingDate: '',
    paket: '',
    package: '',
    odp: '',
    mitra: '',
    symptom: '',
    memo: '',
    tikor: '',
    snOnt: '',
    nikOnt: '',
    stbId: '',
    nikStb: '',
    teknisi: (username || '').replace('@', ''),
  };

  const patterns = {
    channel: /CHANNEL\s*:\s*(.+?)(?=\n|$)/i,
    dateCreated: /DATE\s*CREATED\s*:\s*(.+?)(?=\n|$)/i,
    scOrderNo: /SC\s*ORDER\s*NO\s*:\s*(.+?)(?=\n|$)/i,
    workorder: /WORKORDER\s*:\s*(.+?)(?=\n|$)/i,
    ao: /AO\s*:\s*(.+?)(?=\n|$)/i,
    ncli: /NCLI\s*:\s*(.+?)(?=\n|$)/i,
    serviceNo: /SERVICE\s*NO\s*:\s*(.+?)(?=\n|$)/i,
    address: /ADDRESS\s*:\s*(.+?)(?=\n|$)/i,
    customerName: /CUSTOMER\s*NAME\s*:\s*(.+?)(?=\n|$)/i,
    workzone: /WORKZONE\s*:\s*(.+?)(?=\n|$)/i,
    contactPhone: /CONTACT\s*PHONE\s*:\s*(.+?)(?=\n|$)/i,
    bookingDate: /BOOKING\s*DATE\s*:\s*(.+?)(?=\n|$)/i,
    paket: /PAKET\s*:\s*(.+?)(?=\n|$)/i,
    package: /PACKAGE\s*:\s*(.+?)(?=\n|$)/i,
    odp: /ODP\s*:\s*(.+?)(?=\n|$)/i,
    mitra: /MITRA\s*:\s*(.+?)(?=\n|$)/i,
    symptom: /SYMPTOM\s*:\s*(.+?)(?=\n|$)/i,
    memo: /MEMO\s*:\s*(.+?)(?=\n|$)/i,
    tikor: /TIKOR\s*:\s*(.+?)(?=\n|$)/i,
    snOnt: /SN\s*ONT\s*:\s*(.+?)(?=\n|$)/i,
    nikOnt: /NIK\s*ONT\s*:\s*(.+?)(?=\n|$)/i,
    stbId: /STB\s*ID\s*:\s*(.+?)(?=\n|$)/i,
    nikStb: /NIK\s*STB\s*:\s*(.+?)(?=\n|$)/i,
  };

  for (const [key, pattern] of Object.entries(patterns)) {
    const match = text.match(pattern);
    if (match && match[1]) {
      data[key] = match[1].trim();
    }
  }

  return data;
}

// === BOT SETUP ===
const PORT = process.env.PORT || 3001;
const RAILWAY_STATIC_URL = process.env.RAILWAY_STATIC_URL;
const USE_WEBHOOK = !!RAILWAY_STATIC_URL;

let bot;

if (USE_WEBHOOK) {
  const express = require('express');
  const app = express();
  app.use(express.json());

  bot = new TelegramBot(TOKEN);
  const webhookUrl = `https://${RAILWAY_STATIC_URL}/progres${TOKEN}`;

  bot.setWebHook(webhookUrl).then(() => {
    console.log(`✅ Webhook set: ${webhookUrl}`);
  }).catch(err => {
    console.error('❌ Webhook error:', err.message);
  });

  app.post(`/progres${TOKEN}`, (req, res) => {
    bot.processUpdate(req.body);
    res.sendStatus(200);
  });

  app.get('/', (req, res) => {
    res.send('Bot Progres PSB is running!');
  });

  app.listen(PORT, () => {
    console.log(`✅ Server running on port ${PORT}`);
  });
} else {
  bot = new TelegramBot(TOKEN, { polling: true });
  console.log('✅ Bot running in polling mode');
}

// === MESSAGE HANDLER ===
bot.on('message', async (msg) => {
  const chatId = msg.chat.id;
  const msgId = msg.message_id;
  const text = (msg.text || '').trim();
  const username = msg.from.username || '';

  console.log(`📨 [${username}] ${text.substring(0, 60)}`);

  try {
    // === /UPDATE ===
    if (/^\/UPDATE\b/i.test(text)) {
      const auth = await checkAuthorization(username, ['USER', 'ADMIN']);
      if (!auth.authorized) {
        return sendTelegram(chatId, auth.message, { reply_to_message_id: msgId });
      }

      const inputText = text.replace(/^\/UPDATE\s*/i, '').trim();
      if (!inputText) {
        return sendTelegram(chatId, '❌ Silakan kirim data progres setelah /UPDATE.', { reply_to_message_id: msgId });
      }

      const user = await getUserData(username);
      const parsed = parseProgres(inputText, user, username);

      const required = ['channel', 'scOrderNo', 'serviceNo', 'customerName', 'workzone'];
      const missing = required.filter(f => !parsed[f]);

      if (missing.length > 0) {
        return sendTelegram(chatId, `❌ Field wajib: ${missing.join(', ')}`, { reply_to_message_id: msgId });
      }

      const row = [
        parsed.dateCreated,    // A: DATE CREATED
        parsed.channel,        // B: CHANNEL
        parsed.workorder,      // C: WORKORDER
        parsed.ao,             // D: AO
        parsed.scOrderNo,      // E: SC ORDER NO
        parsed.serviceNo,      // F: SERVICE NO
        parsed.customerName,   // G: CUSTOMER NAME
        parsed.workzone,       // H: WORKZONE
        parsed.contactPhone,   // I: CONTACT PHONE
        parsed.odp,            // J: ODP
        parsed.symptom,        // K: SYMPTOM
        parsed.memo,           // L: MEMO
        parsed.tikor,          // M: TIKOR
        parsed.snOnt,          // N: SN ONT
        parsed.nikOnt,         // O: NIK ONT
        parsed.stbId,          // P: STB ID
        parsed.nikStb,         // Q: NIK STB
        parsed.teknisi,        // R: NAMA TELEGRAM TEKNISI
      ];

      await appendSheetData(PROGRES_SHEET, row);

      let confirmMsg = '✅ Data berhasil disimpan!\n\n';


      return sendTelegram(chatId, confirmMsg, { reply_to_message_id: msgId });
    }

    // === /AKTIVASI ===
    else if (/^\/AKTIVASI\b/i.test(text)) {
      const auth = await checkAuthorization(username, ['USER', 'ADMIN']);
      if (!auth.authorized) {
        return sendTelegram(chatId, auth.message, { reply_to_message_id: msgId });
      }

      try {
        const inputText = text.replace(/^\/AKTIVASI\s*/i, '').trim();
        if (!inputText) {
          return sendTelegram(chatId, '❌ Silakan kirim data aktivasi setelah /AKTIVASI.', { reply_to_message_id: msgId });
        }

        const parsed = parseAktivasi(inputText, username);
        console.log('✅ Parsed aktivasi data:', parsed);

        const required = ['channel', 'customerName', 'serviceNo', 'workzone'];
        const missing = required.filter(f => !parsed[f]);

        if (missing.length > 0) {
          return sendTelegram(chatId, `❌ Field wajib: ${missing.join(', ')}`, { reply_to_message_id: msgId });
        }

        // Tentukan package (gunakan preferensi PACKAGE untuk DIGIPOS, PAKET untuk BS/ES/GS)
        const packageInfo = parsed.package || parsed.paket || '-';

        const row = [
          new Date().toLocaleDateString('id-ID', {
            weekday: 'long',
            day: 'numeric',
            month: 'long',
            year: 'numeric',
            timeZone: 'Asia/Jakarta',
          }),                       // A: DATE CREATED
          parsed.channel,           // B: CHANNEL
          parsed.workorder,         // C: WORKORDER
          parsed.ao,                // D: AO
          parsed.scOrderNo,         // E: SC ORDER NO
          parsed.serviceNo,         // F: SERVICE NO
          parsed.customerName,      // G: CUSTOMER NAME
          parsed.workzone,          // H: WORKZONE
          parsed.contactPhone,      // I: CONTACT PHONE
          parsed.odp,               // J: ODP
          parsed.symptom,           // K: SYMPTOM
          parsed.memo,              // L: MEMO
          parsed.tikor,             // M: TIKOR
          parsed.snOnt,             // N: SN ONT
          parsed.nikOnt,            // O: NIK ONT
          parsed.stbId,             // P: STB ID
          parsed.nikStb,            // Q: NIK STB
          parsed.teknisi,           // R: NAMA TELEGRAM TEKNISI
        ];

        console.log('📝 Row data to append:', row);
        await appendSheetData(PROGRES_SHEET, row);

        let confirmMsg = '✅ Data aktivasi berhasil disimpan!\n\n';;

        return sendTelegram(chatId, confirmMsg, { reply_to_message_id: msgId });
      } catch (aktivasiErr) {
        console.error('❌ /AKTIVASI Error:', aktivasiErr);
        console.error('Stack:', aktivasiErr.stack);
        return sendTelegram(chatId, `❌ Error: ${aktivasiErr.message}`, { reply_to_message_id: msgId });
      }
    }

    // === /today [TEKNISI] ===
    else if (/^\/today\b/i.test(text)) {
      const auth = await checkAuthorization(username, ['ADMIN']);
      if (!auth.authorized) {
        return sendTelegram(chatId, auth.message, { reply_to_message_id: msgId });
      }

      const args = text.replace(/^\/today\s*/i, '').trim();
      if (!args) {
        return sendTelegram(chatId, '❌ Format: /today TEKNISI_ID', { reply_to_message_id: msgId });
      }

      const today = new Date().toLocaleDateString('id-ID', {
        weekday: 'long',
        day: 'numeric',
        month: 'long',
        year: 'numeric',
        timeZone: 'Asia/Jakarta',
      });

      const data = await getSheetData(PROGRES_SHEET);
      let map = {};

      for (let i = 1; i < data.length; i++) {
        const dateCreated = (data[i][0] || '').trim();  // Column A
        if (dateCreated !== today) continue;
        
        const teknisi = (data[i][17] || '-').trim();  // Column R
        if (teknisi !== args) continue;
        
        const symptom = (data[i][10] || '-').trim();  // Column K
        const ao = (data[i][3] || '-').trim();  // Column D (AO)

        if (!map[symptom]) map[symptom] = [];
        map[symptom].push(ao);
      }

      const entries = Object.entries(map)
        .sort((a, b) => b[1].length - a[1].length);

      let totalWO = Object.values(map).reduce((sum, arr) => sum + arr.length, 0);
      let msg = `📋 <b>PROGRES HARI INI - ${args}</b>\n\n`;
      msg += `<b>Total: ${totalWO} WO</b>\n`;
      
      if (entries.length === 0) {
        msg += '<i>Belum ada data untuk hari ini</i>';
      } else {
        entries.forEach((entry) => {
          const [symptom, aos] = entry;
          msg += `   • <b>${symptom}: ${aos.length}</b>\n`;
          aos.forEach(ao => {
            msg += `${ao}\n`;
          });
          msg += '\n';
        });
      }

      return sendTelegram(chatId, msg, { reply_to_message_id: msgId });
    }

    // === /all [TEKNISI] ===
    else if (/^\/all\b/i.test(text)) {
      const auth = await checkAuthorization(username, ['ADMIN']);
      if (!auth.authorized) {
        return sendTelegram(chatId, auth.message, { reply_to_message_id: msgId });
      }

      const args = text.replace(/^\/all\s*/i, '').trim();
      if (!args) {
        return sendTelegram(chatId, '❌ Format: /all TEKNISI_ID', { reply_to_message_id: msgId });
      }

      const data = await getSheetData(PROGRES_SHEET);
      let map = {};

      for (let i = 1; i < data.length; i++) {
        const teknisi = (data[i][17] || '-').trim();  // Column R
        if (teknisi !== args) continue;
        
        const symptom = (data[i][10] || '-').trim();  // Column K
        const ao = (data[i][3] || '-').trim();  // Column D (AO)

        if (!map[symptom]) map[symptom] = [];
        map[symptom].push(ao);
      }

      const entries = Object.entries(map)
        .sort((a, b) => b[1].length - a[1].length);

      let totalWO = Object.values(map).reduce((sum, arr) => sum + arr.length, 0);
      let msg = `📋 <b>SELURUH PROGRES - ${args}</b>\n\n`;
      msg += `<b>Total: ${totalWO} WO</b>\n`;
      
      if (entries.length === 0) {
        msg += '<i>Belum ada data</i>';
      } else {
        entries.forEach((entry) => {
          const [symptom, aos] = entry;
          msg += `   • <b>${symptom}: ${aos.length}</b>\n`;
          aos.forEach(ao => {
            msg += `${ao}\n`;
          });
          msg += '\n';
        });
      }

      return sendTelegram(chatId, msg, { reply_to_message_id: msgId });
    }

    // === /progres ===
    else if (/^\/progres\b/i.test(text)) {
      const auth = await checkAuthorization(username, ['ADMIN']);
      if (!auth.authorized) {
        return sendTelegram(chatId, auth.message, { reply_to_message_id: msgId });
      }

      const today = new Date().toLocaleDateString('id-ID', {
        weekday: 'long',
        day: 'numeric',
        month: 'long',
        year: 'numeric',
        timeZone: 'Asia/Jakarta',
      });

      const data = await getSheetData(PROGRES_SHEET);
      let map = {};

      for (let i = 1; i < data.length; i++) {
        const dateCreated = (data[i][0] || '').trim();  // Column A (index 0)
        if (dateCreated !== today) continue;  // Filter hanya hari ini
        
        const teknisi = (data[i][17] || '-').trim();  // Column R (index 17)
        const symptom = (data[i][10] || '-').trim();  // Column K (index 10)

        if (!map[teknisi]) map[teknisi] = { total: 0 };
        map[teknisi].total++;
        map[teknisi][symptom] = (map[teknisi][symptom] || 0) + 1;
      }

      const entries = Object.entries(map)
        .sort((a, b) => b[1].total - a[1].total);

      let msg = `📊 <b>LAPORAN TEKNISI - HARI INI</b>\n<b>${today}</b>\n\n`;
      
      if (entries.length === 0) {
        msg += '<i>Belum ada data untuk hari ini</i>';
      } else {
        entries.forEach((entry) => {
          const [teknisi, counts] = entry;
          msg += `🔸 <b>${teknisi}</b>\n`;
          msg += `   <b>Total:</b> ${counts.total} WO\n`;
          
          const symptoms = Object.entries(counts)
            .filter(([k]) => k !== 'total')
            .sort((a, b) => b[1] - a[1]);
          
          symptoms.forEach(([symptomName, count]) => {
            msg += `   • ${symptomName}: ${count}\n`;
          });
          msg += '\n';
        });
      }

      return sendTelegram(chatId, msg, { reply_to_message_id: msgId });
    }

    // === /allprogres ===
    else if (/^\/allprogres\b/i.test(text)) {
      const auth = await checkAuthorization(username, ['ADMIN']);
      if (!auth.authorized) {
        return sendTelegram(chatId, auth.message, { reply_to_message_id: msgId });
      }

      const data = await getSheetData(PROGRES_SHEET);
      let map = {};

      for (let i = 1; i < data.length; i++) {
        const teknisi = (data[i][17] || '-').trim();  // Column R (index 17)
        const symptom = (data[i][10] || '-').trim();  // Column K (index 10)

        if (!map[teknisi]) map[teknisi] = { total: 0 };
        map[teknisi].total++;
        map[teknisi][symptom] = (map[teknisi][symptom] || 0) + 1;
      }

      const entries = Object.entries(map)
        .sort((a, b) => b[1].total - a[1].total);

      let msg = '📊 <b>LAPORAN TEKNISI - KESELURUHAN</b>\n\n';
      
      if (entries.length === 0) {
        msg += '<i>Belum ada data</i>';
      } else {
        entries.forEach((entry) => {
          const [teknisi, counts] = entry;
          msg += `🔸 <b>${teknisi}</b>\n`;
          msg += `   <b>Total:</b> ${counts.total} WO\n`;
          
          const symptoms = Object.entries(counts)
            .filter(([k]) => k !== 'total')
            .sort((a, b) => b[1] - a[1]);
          
          symptoms.forEach(([symptomName, count]) => {
            msg += `   • ${symptomName}: ${count}\n`;
          });
          msg += '\n';
        });
      }

      return sendTelegram(chatId, msg, { reply_to_message_id: msgId });
    }

    // === /cek ===
    else if (/^\/cek\b/i.test(text)) {
      const auth = await checkAuthorization(username, ['ADMIN']);
      if (!auth.authorized) {
        return sendTelegram(chatId, auth.message, { reply_to_message_id: msgId });
      }

      const today = new Date().toLocaleDateString('id-ID', {
        weekday: 'long',
        day: 'numeric',
        month: 'long',
        year: 'numeric',
        timeZone: 'Asia/Jakarta',
      });

      const data = await getSheetData(PROGRES_SHEET);
      let map = {};

      for (let i = 1; i < data.length; i++) {
        const dateCreated = (data[i][0] || '').trim();  // Column A (index 0)
        if (dateCreated !== today) continue;  // Filter hanya hari ini
        
        const workzone = (data[i][7] || '-').trim();   // Column H (index 7)
        const symptom = (data[i][10] || '-').trim();   // Column K (index 10)

        if (!map[workzone]) map[workzone] = { total: 0 };
        map[workzone].total++;
        map[workzone][symptom] = (map[workzone][symptom] || 0) + 1;
      }

      const entries = Object.entries(map)
        .sort((a, b) => b[1].total - a[1].total);

      let msg = `📍 <b>REKAP WORKZONE - HARI INI</b>\n<b>${today}</b>\n\n`;
      
      if (entries.length === 0) {
        msg += '<i>Belum ada data untuk hari ini</i>';
      } else {
        entries.forEach((entry) => {
          const [workzone, counts] = entry;
          msg += `🔸 <b>${workzone}</b>\n`;
          msg += `   <b>Total:</b> ${counts.total} WO\n`;
          
          const symptoms = Object.entries(counts)
            .filter(([k]) => k !== 'total')
            .sort((a, b) => b[1] - a[1]);
          
          symptoms.forEach(([symptomName, count]) => {
            msg += `   • ${symptomName}: ${count}\n`;
          });
          msg += '\n';
        });
      }

      return sendTelegram(chatId, msg, { reply_to_message_id: msgId });
    }

    // === /allcek ===
    else if (/^\/allcek\b/i.test(text)) {
      const auth = await checkAuthorization(username, ['ADMIN']);
      if (!auth.authorized) {
        return sendTelegram(chatId, auth.message, { reply_to_message_id: msgId });
      }

      const data = await getSheetData(PROGRES_SHEET);
      let map = {};

      for (let i = 1; i < data.length; i++) {
        const workzone = (data[i][7] || '-').trim();   // Column H (index 7)
        const symptom = (data[i][10] || '-').trim();   // Column K (index 10)

        if (!map[workzone]) map[workzone] = { total: 0 };
        map[workzone].total++;
        map[workzone][symptom] = (map[workzone][symptom] || 0) + 1;
      }

      const entries = Object.entries(map)
        .sort((a, b) => b[1].total - a[1].total);

      let msg = '📍 <b>REKAP WORKZONE - KESELURUHAN</b>\n\n';
      
      if (entries.length === 0) {
        msg += '<i>Belum ada data</i>';
      } else {
        entries.forEach((entry) => {
          const [workzone, counts] = entry;
          msg += `🔸 <b>${workzone}</b>\n`;
          msg += `   <b>Total:</b> ${counts.total} WO\n`;
          
          const symptoms = Object.entries(counts)
            .filter(([k]) => k !== 'total')
            .sort((a, b) => b[1] - a[1]);
          
          symptoms.forEach(([symptomName, count]) => {
            msg += `   • ${symptomName}: ${count}\n`;
          });
          msg += '\n';
        });
      }

      return sendTelegram(chatId, msg, { reply_to_message_id: msgId });
    }

    // === /help ===
    else if (/^\/(help|start)\b/i.test(text)) {
      const auth = await checkAuthorization(username);
      if (!auth.authorized) {
        return sendTelegram(chatId, auth.message, { reply_to_message_id: msgId });
      }

      const helpMsg = `🤖 Bot Progres PSB

Commands:
/UPDATE - Input progres (di group)
/AKTIVASI - Input data aktivasi (di group)
/today TEKNISI_ID - Progres hari ini (teknisi)
/all TEKNISI_ID - Seluruh progres (teknisi)
/progres - Laporan teknisi HARI INI
/allprogres - Laporan teknisi KESELURUHAN
/cek - Rekap workzone HARI INI
/allcek - Rekap workzone KESELURUHAN
/help - Bantuan

Contoh:
/today FH_ABDULLAH_16891190
/all FH_ABDULLAH_16891190
/progres
/allprogres
/cek
/allcek`;

      return sendTelegram(chatId, helpMsg, { reply_to_message_id: msgId });
    }

  } catch (err) {
    console.error('Error:', err.message);
    sendTelegram(chatId, '❌ Terjadi kesalahan sistem.', { reply_to_message_id: msgId });
  }
});

process.on('unhandledRejection', (reason) => {
  console.error('Error:', reason);
});

console.log('\n🚀 Bot Progres PSB started!');
console.log(`Mode: ${USE_WEBHOOK ? 'Webhook' : 'Polling'}`);

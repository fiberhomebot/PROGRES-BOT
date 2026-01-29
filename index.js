require('dotenv').config();
const { Telegraf } = require('telegraf');
const { GoogleSheets } = require('./google_sheets');

const BOT_TOKEN = process.env.BOT_TOKEN;
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const SHEET_NAME = process.env.SHEET_NAME || 'Sheet1';
// forwarding removed; no FORWARD_CHAT_ID used

if (!BOT_TOKEN || !SPREADSHEET_ID) {
  console.error('Missing BOT_TOKEN or SPREADSHEET_ID in environment');
  process.exit(1);
}

if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
  console.error('Missing GOOGLE_SERVICE_ACCOUNT_JSON in environment. Do NOT commit service account JSON to git; set it as an env var in Railway.');
  process.exit(1);
}

const bot = new Telegraf(BOT_TOKEN);
const sheets = new GoogleSheets({ spreadsheetId: SPREADSHEET_ID, sheetName: SHEET_NAME });

function parseUpdateMessage(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  const result = {};
  for (const line of lines) {
    // skip the command line if present
    if (line.toUpperCase().startsWith('/UPDATE')) continue;
    const parts = line.split(':');
    if (parts.length >= 2) {
      const key = parts.shift().trim().toUpperCase().replace(/\s+/g, ' ');
      const value = parts.join(':').trim();
      result[key] = value;
    }
  }
  return result;
}

function formatSummary(obj) {
  let out = '';
  for (const k of Object.keys(obj)) {
    out += `*${k}*: ${obj[k]}\n`;
  }
  return out;
}

bot.command('update', async (ctx) => {
  try {
    const text = ctx.message.text || '';
    const parsed = parseUpdateMessage(text);
    // attach telegram metadata
    parsed['TG_CHAT_ID'] = String(ctx.chat.id);
    parsed['TG_MSG_ID'] = String(ctx.message.message_id);
    parsed['TG_USERNAME'] = ctx.from.username || `${ctx.from.first_name || ''} ${ctx.from.last_name || ''}`.trim();
    parsed['TANGGAL'] = parsed['DATE CREATED'] || new Date().toLocaleString('id-ID');

    await sheets.appendRow(parsed);

    await ctx.reply('✅ Data berhasil diproses dan disimpan ke sheet.');

    // forwarding feature removed (was sending summary to configured group)
  } catch (err) {
    console.error('Error handling /update', err);
    await ctx.reply('❌ Terjadi kesalahan saat menyimpan data.');
  }
});

// Start periodic checker for CATATAN KORLAP -> reply to user message id
async function checkNotesAndReply() {
  try {
    const rows = await sheets.getAllRows();
    if (!rows || rows.length === 0) return;
    const headers = rows[0];
    const dataRows = rows.slice(1);
    const idxChat = headers.findIndex(h => (h||'').toUpperCase() === 'TG_CHAT_ID');
    const idxMsg = headers.findIndex(h => (h||'').toUpperCase() === 'TG_MSG_ID');
    const idxNote = headers.findIndex(h => (h||'').toUpperCase() === 'CATATAN KORLAP');

    if (idxNote === -1 || idxChat === -1 || idxMsg === -1) return; // required columns

    // ensure REPLIED column exists and get its index
    const idxReplied = await sheets.ensureHeader('REPLIED');

    for (let i = 0; i < dataRows.length; i++) {
      const row = dataRows[i];
      const note = row[idxNote] || '';
      const replied = (idxReplied !== -1) ? (row[idxReplied] || '') : '';
      if (note && !replied) {
        const chatId = row[idxChat];
        const msgId = Number(row[idxMsg]);
        try {
          await bot.telegram.sendMessage(chatId, `Balasan dari Korlap:\n${note}`, { reply_to_message_id: msgId });
          // mark replied in sheet
          if (idxReplied !== -1) {
            const sheetRowIndex = i + 2; // because header row + 1-based
            await sheets.updateCell(sheetRowIndex, idxReplied + 1, new Date().toLocaleString('id-ID'));
          }
        } catch (err) {
          console.error('Failed sending reply for row', i + 2, err);
        }
      }
    }
  } catch (err) {
    console.error('Error in checkNotesAndReply', err);
  }
}

// poll every 60 seconds
setInterval(checkNotesAndReply, Number(process.env.POLL_INTERVAL_MS) || 60000);

bot.launch().then(() => {
  console.log('Bot started');
});

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));

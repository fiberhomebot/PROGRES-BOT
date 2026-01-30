// Shim to accept alternate env var names and start the main bot
process.env.TELEGRAM_TOKEN = process.env.TELEGRAM_TOKEN || process.env.BOT_TOKEN;
process.env.SHEET_ID = process.env.SHEET_ID || process.env.SPREADSHEET_ID;

require('./bot-progres-psb.js');

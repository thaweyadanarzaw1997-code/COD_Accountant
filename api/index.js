const TelegramBot = require('node-telegram-bot-api');
const { google } = require('googleapis');
const XLSX = require('xlsx');
const axios = require('axios');

// Environment Variables များမှ ယူမည်
const token = process.env.TELEGRAM_BOT_TOKEN;
const SPREADSHEET_ID = process.env.GOOGLE_SHEET_ID;

const bot = new TelegramBot(token);

// Google Auth Setup
const auth = new google.auth.GoogleAuth({
  credentials: {
    client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
  },
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const sheets = google.sheets({ version: 'v4', auth });

module.exports = async (req, res) => {
  if (req.method === 'POST') {
    const update = req.body;

    if (update.message && update.message.document) {
      const doc = update.message.document;
      const fileName = doc.file_name;

      // Excel ဖိုင် ဟုတ်မဟုတ် စစ်မည် (.xlsx သို့မဟုတ် .xls)
      if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
        try {
          // Telegram ကနေ ဖိုင်လင့်ခ် တောင်းမည်
          const fileLink = await bot.getFileLink(doc.file_id);
          
          // ဖိုင်ကို Download ဆွဲမည်
          const response = await axios.get(fileLink, { responseType: 'arraybuffer' });
          const workbook = XLSX.read(response.data, { type: 'buffer' });
          
          // ပထမဆုံး Sheet ကို ယူမည်
          const sheetName = workbook.SheetNames[0];
          const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

          const rows = sheetData.slice(1);

          if (rows.length === 0) {
            await bot.sendMessage(update.message.chat.id, 'ပို့လာသော Excel ဖိုင်ထဲတွင် ဒေတာ မပါရှိပါ။');
            return res.status(200).send('OK');
          }

          // ၁။ Google Sheet (Sheet3) ထဲမှာ ရှိပြီးသား Waybill (E Column - Index 4) တွေကို အရင်ဆွဲထုတ်မည်
          const existingSheet = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Sheet3!E:E', // Sheet3 ရဲ့ E Column ကို စစ်မည်
          });

          const existingWaybills = new Set();
          if (existingSheet.data.values) {
            existingSheet.data.values.forEach(row => {
              if (row[0]) existingWaybills.add(row[0].toString().trim());
            });
          }

          // ၂။ မပါသေးတဲ့ အသစ်တွေကိုပဲ ရွေးထုတ်မည်
          const newRowsToAppend = [];
          for (const row of rows) {
            const waybillVal = row[4]; // E Column သည် Index 4 ဖြစ်သည်
            if (waybillVal) {
              const cleanWaybill = waybillVal.toString().trim();
              if (!existingWaybills.has(cleanWaybill)) {
                newRowsToAppend.push(row);
                existingWaybills.add(cleanWaybill);
              }
            }
          }

          if (newRowsToAppend.length > 0) {
            // Google Sheet (Sheet3) ထဲသို့ Append လုပ်မည်
            await sheets.spreadsheets.values.append({
              spreadsheetId: SPREADSHEET_ID,
              range: 'Sheet3!A:O', // Sheet3 ထဲသို့ A ကနေ O အထိ ဖြည့်မည်
              valueInputOption: 'USER_ENTERED',
              resource: {
                values: newRowsToAppend,
              },
            });

            await bot.sendMessage(
              update.message.chat.id, 
              `✅ အောင်မြင်ပါသည်။ Sheet3 သို့ အသစ်ပါလာသော Waybill စုစုပေါင်း (${newRowsToAppend.length}) ခု ကို ဖြည့်ပြီးပါပြီ။`
            );
          } else {
            await bot.sendMessage(
              update.message.chat.id, 
              `ℹ️ ပို့လိုက်သော Excel ဖိုင်ထဲမှ Waybill အားလုံးသည် Sheet3 ထဲတွင် ရှိနှင့်ပြီးသားဖြစ်သောကြောင့် ဘာမှအသစ်ထည့်စရာမလိုပါ။`
            );
          }

        } catch (error) {
          console.error(error);
          await bot.sendMessage(update.message.chat.id, '❌ Error ဖြစ်သွားပါသည်: ' + error.message);
        }
      } else {
        await bot.sendMessage(update.message.chat.id, 'ကျေးဇူးပြု၍ .xlsx ပုံစံ Excel ဖိုင်ကိုသာ ပို့ပေးပါ။');
      }
    }

    res.status(200).send('OK');
  } else {
    res.status(200).send('Telegram Bot Webhook is running!');
  }
};

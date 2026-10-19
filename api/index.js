const TelegramBot = require('node-telegram-bot-api');
const { google } = require('googleapis');
const XLSX = require('xlsx');
const axios = require('axios');

const token = process.env.TELEGRAM_BOT_TOKEN;
const SPREADSHEET_ID = process.env.GOOGLE_SHEET_ID;

const bot = new TelegramBot(token);

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

    try {
      // 1. Telegram ကနေ Document (Excel) ပို့လာရင်
      if (update.message && update.message.document) {
        const doc = update.message.document;
        const fileName = doc.file_name;

        if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
          const fileLink = await bot.getFileLink(doc.file_id);
          const response = await axios.get(fileLink, { responseType: 'arraybuffer' });
          const workbook = XLSX.read(response.data, { type: 'buffer' });
          
          const sheetName = workbook.SheetNames[0];
          const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
          const rows = sheetData.slice(1);

          if (rows.length === 0) {
            await bot.sendMessage(update.message.chat.id, 'ပို့လာသော Excel ဖိုင်ထဲတွင် ဒေတာ မပါရှိပါ။');
            return res.status(200).send('OK');
          }

          const existingSheet = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Sheet3!E:E',
          });

          const existingWaybills = new Set();
          if (existingSheet.data.values) {
            existingSheet.data.values.forEach(row => {
              if (row[0]) existingWaybills.add(row[0].toString().trim());
            });
          }

          const newRowsToAppend = [];
          for (const row of rows) {
            const waybillVal = row[4]; // E Column
            if (waybillVal) {
              const cleanWaybill = waybillVal.toString().trim();
              if (!existingWaybills.has(cleanWaybill)) {
                newRowsToAppend.push(row);
                existingWaybills.add(cleanWaybill);
              }
            }
          }

          if (newRowsToAppend.length > 0) {
            await sheets.spreadsheets.values.append({
              spreadsheetId: SPREADSHEET_ID,
              range: 'Sheet3!A:O',
              valueInputOption: 'USER_ENTERED',
              resource: { values: newRowsToAppend },
            });

            await bot.sendMessage(
              update.message.chat.id, 
              `✅ Sheet3 သို့ အသစ်ပါလာသော Waybill (${newRowsToAppend.length}) ခု ကို ဖြည့်ပြီးပါပြီ။`
            );
          } else {
            await bot.sendMessage(
              update.message.chat.id, 
              `ℹ️ Waybill အားလုံးသည် Sheet3 ထဲတွင် ရှိနှင့်ပြီးသား ဖြစ်ပါသည်။`
            );
          }
        } else {
          await bot.sendMessage(update.message.chat.id, 'ကျေးဇူးပြု၍ .xlsx ပုံစံ Excel ဖိုင်ကိုသာ ပို့ပေးပါ။');
        }
      }

      // 2. Text Message များအတွက် (ဥပမာ: Order အသစ်ပြရန် သို့မဟုတ် Waybill ရိုက်ထည့်ရန်)
      if (update.message && update.message.text) {
        const text = update.message.text.trim();
        const chatId = update.message.chat.id;

        if (text === '/orders') {
          // Sheet2 ထဲက Order အသစ်များကို ဆွဲထုတ်ပြပြီး In Progress လုပ်ရန် Button ပြမည်
          const sheet2Data = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Sheet2!A:E', // လိုအပ်သော Column များကို ယူရန်
          });

          const rows = sheet2Data.data.values;
          if (!rows || rows.length <= 1) {
            await bot.sendMessage(chatId, 'Sheet2 ထဲတွင် Order အသစ်များ မရှိသေးပါ။');
          } else {
            // ဥပမာအနေနဲ့ နောက်ဆုံး Order များကို ပြသပေးခြင်း
            let msg = `📦 **Sheet2 မှ Order စာရင်းများ:**\n\n`;
            rows.slice(-5).forEach((r, idx) => {
              msg += `${idx + 1}. Customer: ${r[3] || '-'} | Waybill: ${r[4] || '-'}\n`;
            });

            await bot.sendMessage(chatId, msg, {
              parse_mode: 'Markdown',
              reply_markup: {
                inline_keyboard: [
                  [{ text: '🔄 In Progress ပြောင်းရန်', callback_data: 'prompt_waybill' }]
                ]
              }
            });
          }
        } else if (text.startsWith('/inprogress')) {
          // ပုံစံ: /inprogress <Waybill_Number>
          const parts = text.split(' ');
          if (parts.length < 2) {
            await bot.sendMessage(chatId, 'ကျေးဇူးပြု၍ Waybill နံပါတ် ထည့်ပေးပါ။ ဥပမာ: /inprogress ABE992994');
          } else {
            const targetWaybill = parts[1].trim();
            
            // Sheet2 ထဲတွင် အဆိုပါ Waybill ကို ရှာမည် (ဥပမာ E Column ထဲတွင် ရှိသည်ဟု ယူဆပါက)
            const sheet2All = await sheets.spreadsheets.values.get({
              spreadsheetId: SPREADSHEET_ID,
              range: 'Sheet2!A:Z',
            });

            const allRows = sheet2All.data.values;
            let rowIndex = -1;

            if (allRows) {
              for (let i = 0; i < allRows.length; i++) {
                if (allRows[i][4] && allRows[i][4].toString().trim() === targetWaybill) {
                  rowIndex = i + 1; // Google Sheet Row index သည် 1 မှ စသည်။
                  break;
                }
              }
            }

            if (rowIndex !== -1) {
              // L Column (Index 11 -> 12th column) တွင် Waybill ထည့်မည်၊ M Column (Index 12 -> 13th column) တွင် "In Progress" ထည့်မည်
              await sheets.spreadsheets.values.update({
                spreadsheetId: SPREADSHEET_ID,
                range: `Sheet2!L${rowIndex}:M${rowIndex}`,
                valueInputOption: 'USER_ENTERED',
                resource: {
                  values: [[targetWaybill, 'In Progress']]
                }
              });

              await bot.sendMessage(chatId, `✅ Waybill (${targetWaybill}) ကို L Column တွင်ထည့်ပြီး M Column ကို 'In Progress' သို့ အောင်မြင်စွာ ပြောင်းပြီးပါပြီ။`);
            } else {
              await bot.sendMessage(chatId, `❌ Sheet2 ထဲတွင် Waybill (${targetWaybill}) ကို မတွေ့ရှိပါ။`);
            }
          }
        }
      }

      // 3. Telegram Inline Button များကို နှိပ်သည့်အခါ (Callback Query)
      if (update.callback_query) {
        const callbackQuery = update.callback_query;
        const chatId = callbackQuery.message.chat.id;

        if (callbackQuery.data === 'prompt_waybill') {
          await bot.sendMessage(chatId, '✏️ In Progress လုပ်လိုသော Waybill နံပါတ်ကို အောက်ပါ ပုံစံအတိုင်း ပို့ပေးပါ:\n\n`/inprogress <Waybill_Number>`');
          await bot.answerCallbackQuery(callbackQuery.id);
        }
      }

    } catch (error) {
      console.error(error);
    }
  }

  res.status(200).send('OK');
};

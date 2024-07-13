const TelegramBot = require('node-telegram-bot-api');
const mongoose = require('mongoose');
const ExcelJS = require('exceljs');
const schedule = require('node-schedule');
const moment = require('moment-timezone');

// Инициализация бота
const bot = new TelegramBot('7351145568:AAFh0GarhCu5MdygEaR5_dkxqydYhfu5Fa0', { polling: true }); // Замените YOUR_BOT_TOKEN на ваш токен
const groupId = '@GFDSAIISRB_grafik';
const password = 'banana'; // Замените на необходимый пароль

// Подключение к MongoDB
mongoose.connect('mongodb://localhost:27017/telegramBotDB', { useNewUrlParser: true, useUnifiedTopology: true });

const userSchema = new mongoose.Schema({
  telegramId: String,
  username: String,
  phoneNumber: String,
  fullName: String,
  degree: String,
  organization: String,
});

const User = mongoose.model('User', userSchema);

// Обработчик команды /start
bot.onText(/\/start/, (msg) => {
  const chatId = msg.chat.id;
  bot.sendMessage(chatId, 'Iltimos parolni kiriting:');
  
  // Ожидание ввода пароля
  bot.once('message', async (msg) => {
    if (msg.text === password) {
      const user = await User.findOne({ telegramId: msg.from.id });
      if (user) {
        bot.sendMessage(chatId, `Siz oldin ro'yhatdan o'tgan ekansiz. Bugun qaysi tashkilotga borasiz ?`);
        
        // Ожидание ввода организации для зарегистрированного пользователя
        bot.once('message', async (msg) => {
          user.organization = msg.text;
          await user.save();
          bot.sendMessage(chatId, `Tashkilot yangilandi va rahbariyatga yuborildi`);
        });
      } else {
        bot.sendMessage(chatId, `Iltimos, telefon raqamingizni yozing:`);
        
        // Ожидание ввода номера телефона
        bot.once('message', async (msg) => {
          const phoneNumber = msg.text;
          bot.sendMessage(chatId, `Familiya Isim Sharifingizni to'liq yozing:`);
          
          // Ожидание ввода полного имени
          bot.once('message', async (msg) => {
            const fullName = msg.text;
            bot.sendMessage(chatId, `Lavozimingizni to'liq yozing`);
            
            // Ожидание ввода должности
            bot.once('message', async (msg) => {
              const degree = msg.text;
              const newUser = new User({
                telegramId: msg.from.id,
                username: msg.from.username,
                phoneNumber,
                fullName,
                degree,
                organization: '',
              });
              await newUser.save();
              bot.sendMessage(chatId, `Tabriklayman siz ro'yhatdan to'liq o'ttingiz. Endi bugun qaysi tashkilotda bo'lishingizni yozing?`);
              
              // Ожидание ввода организации для нового пользователя
              bot.once('message', async (msg) => {
                newUser.organization = msg.text;
                await newUser.save();
                bot.sendMessage(chatId, `Ro'yhatdan o'tish yakunlandi`);
              });
            });
          });
        });
      }
    } else {
      bot.sendMessage(chatId, `Parol noto'gri terildi. Qayta urunib ko'ring.`);
    }
  });
});

// Функция для генерации и отправки отчета
async function generateAndSendReport() {
  const users = await User.find();
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('User Data');
  worksheet.columns = [
    { header: 'F.I.Sh', key: 'fullName', width: 40 },
    { header: 'Telefon raqami', key: 'phoneNumber', width: 20 },
    { header: 'Lavozimi', key: 'degree', width: 30 },
    { header: 'Tashkilot nomi', key: 'organization', width: 110 },
  ];

  users.forEach(user => {
    worksheet.addRow({
      fullName: user.fullName,
      phoneNumber: user.phoneNumber,
      degree: user.degree,
      organization: user.organization || 'Xodim xozircha tomondan Malumot kiritilmagan',
    }).eachCell((cell, colNumber) => {
      if (colNumber === 4 && user.organization === '') {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFF6666' } // Красный цвет
        };
        cell.font = { size: 16, bold: true };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.value = 'Xodim xozircha tomondan Malumot kiritilmagan';
      } else if (colNumber === 4) {
        cell.font = { name: 'Times New Roman', size: 16 }; // Для пустых ячеек в колонке использовать другой шрифт и размер
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }); 
  });

  // Применение стиля для первой строки отдельно
  worksheet.getRow(1).eachCell((cell) => {
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.font = { size: 14, bold: true };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFCCFFCC' } // Светло-зеленый цвет
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } }
    }
     // Применение выравнивания текста слева для всех столбцов, кроме первого
    if (cell.col !== 1) {
      cell.alignment = { horizontal: 'center' };
  }
  });

  // Применение стиля для всех заполненных строк
  // worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
  //   if (rowNumber > 1) {
  //     row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
  //       cell.border = {
  //         top: { style: 'thin', color: { argb: 'FF000000' } },
  //         left: { style: 'thin', color: { argb: 'FF000000' } },
  //         bottom: { style: 'thin', color: { argb: 'FF000000' } },
  //         right: { style: 'thin', color: { argb: 'FF000000' } }
  //       };
  //       if (colNumber <= 3) {
  //         cell.font = { name: 'Times New Roman', size: 14 };
  //         cell.alignment = { vertical: 'middle', horizontal: 'center' };
  //       }
  //     });
  //   }
  // });
  // Применение стиля для всех заполненных строк, исключая первую колонку, начиная со второй строки
worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
  if (rowNumber > 1) {
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      cell.border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
      };
      if (colNumber !== 1) { // Исключаем первую колонку
        cell.font = { name: 'Times New Roman', size: 14 };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    });
  }
});


  // Генерация файла Excel и отправка в группу
  const buffer = await workbook.xlsx.writeBuffer();
  bot.sendDocument(groupId, buffer, {}, { filename: 'UserData.xlsx', contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

// Обработчик всех сообщений, кроме /start и пароля
bot.on('message', async (msg) => {
  if (msg.text !== '/start' && msg.text !== password) {
    const user = await User.findOne({ telegramId: msg.from.id });
    if (user) {
      user.organization = msg.text;
      await user.save();
      bot.sendMessage(msg.chat.id, 'Tashkilot yangilandi.');
      generateAndSendReport();
    }
  }
});

// Запуск генерации отчета каждый день в 10:00 и каждую минуту для отладки
setInterval(() => {
  const now = new Date();
  if (now.getHours() === 10 && now.getMinutes() === 0) {
    generateAndSendReport();
  }
}, 60000);

// Очистка поля organization ежедневно в 18:30 по Ташкенту
schedule.scheduleJob({ hour: 18, minute: 30, tz: 'Asia/Tashkent' }, async () => {
  await User.updateMany({}, { $set: { organization: '' } });
  console.log('Поле organization очищено для всех пользователей');
});

// Обработка ошибок
process.on('unhandledRejection', error => {
  console.error('Unhandled promise rejection:', error);
});

bot.on('polling_error', error => {
  console.error('Polling error:', error);
});

console.log('Бот успешно запущен и готов к работе.');

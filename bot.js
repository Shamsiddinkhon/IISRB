const TelegramBot = require('node-telegram-bot-api');
const mongoose = require('mongoose');
const ExcelJS = require('exceljs');
const schedule = require('node-schedule'); // Импортируем библиотеку node-schedule
const moment = require('moment-timezone');

// Инициализация бота
const bot = new TelegramBot('7351145568:AAFh0GarhCu5MdygEaR5_dkxqydYhfu5Fa0', { polling: true });
const groupId = '@GFDSAIISRB_grafik';
const password = 'banana'; // Замените на необходимый пароль

// Подключение к MongoDB
mongoose.connect('mongodb://localhost:27017/telegramBotDB');

const userSchema = new mongoose.Schema({
  telegramId: String,
  username: String,
  phoneNumber: String,
  fullName: String,
  degree: String,
  organization: String,
});

const User = mongoose.model('User', userSchema);

bot.onText(/\/start/, (msg) => {
  const chatId = msg.chat.id;
  bot.sendMessage(chatId, 'Iltimos parolni kiriting:');
  bot.once('message', async (msg) => {
    if (msg.text === password) {
      const user = await User.findOne({ telegramId: msg.from.id });
      if (user) {
        bot.sendMessage(chatId, `Siz oldin ro'yhatdan o'tgan ekansiz. Bugun qaysi tashkilotga borasiz ?`);
        bot.once('message', async (msg) => {
          user.organization = msg.text;
          await user.save();
          bot.sendMessage(chatId, 'Tashkilot yangilandi va rahbariyatga yuborildi');
        });
      } else {
        bot.sendMessage(chatId, 'Iltimos, telefon raqamingizni yozing:');
        bot.once('message', async (msg) => {
          const phoneNumber = msg.text;
          bot.sendMessage(chatId, `Familiya Isim Sharifingizni to'liq yozing:`);
          bot.once('message', async (msg) => {
            const fullName = msg.text;
            bot.sendMessage(chatId, `Lavozimingizni to'liq yozing`);
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
              bot.once('message', async (msg) => {
                newUser.organization = msg.text;
                await newUser.save();
                bot.sendMessage(chatId, `Ro'yhatdan o'tish yakunlandi`);//Регистрация завершена.
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

async function generateAndSendReport() {
  const users = await User.find();
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('User Data');
  worksheet.columns = [
    { header: 'F.I.SH', key: 'fullName', width: 40 },
    { header: 'Telefon raqami', key: 'phoneNumber', width: 20 },
    { header: 'Lavozimi', key: 'degree', width: 30 },
    { header: 'Tashkilot nomi', key: 'organization', width: 100 },
  ];

  users.forEach(user => {
    worksheet.addRow({
      fullName: user.fullName,
      phoneNumber: user.phoneNumber,
      degree: user.degree,
      organization: user.organization || `Xozirda xodim tomondan malumot berilmagan`,
    }).eachCell((cell, colNumber) => {
      if (colNumber === 4 && user.organization === '') {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFF0000' } // Red color
        };
        cell.value = `Xozirda xodim tomondan malumot berilmagan`;
      }
    });
  });

  const buffer = await workbook.xlsx.writeBuffer();
  
  const userName = '@' + users[0].username; // Пример взятия userName первого пользователя
  bot.sendDocument(groupId, buffer, { caption: `Grafikga yangi tashkilot kiritildi!` }, { filename: 'UserData.xlsx', contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

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

setInterval(() => {
  const now = new Date();
  if (now.getHours() === 10 && now.getMinutes() === 0) {
    generateAndSendReport();
  }
}, 60000);

// Очистка поля organization ежедневно в 18:30 по Ташкенту
const tashkentTime = moment.tz('Asia/Tashkent');
const clearSchedule = schedule.scheduleJob({ hour: 18, minute: 30, tz: 'Asia/Tashkent' }, async () => {
  await User.updateMany({}, { $set: { organization: '' } });
  console.log('Поле organization очищено для всех пользователей');
});

// // Планировщик для очистки поля organization каждый день в 18:30 по Ташкенту
// const tashkentOffset = 5; // Часовой пояс Ташкента (UTC+5)
// const clearJob = schedule.scheduleJob({ hour: 13, minute: 30, tz: 'Etc/UTC' }, async () => {
//   try {
//     await User.updateMany({}, { $set: { organization: '' } });
//     console.log('Поля organization очищены');
//   } catch (error) {
//     console.error('Ошибка при очистке поля organization:', error);
//   }
// });

// Планировщик для отправки напоминания пользователям каждый будний день в 8:30 по Ташкенту
const reminderJob = schedule.scheduleJob('30 3 * * 1-5', async () => {
  try {
    const users = await User.find();
    users.forEach(user => {
      bot.sendMessage(user.telegramId, `Assalomu aleykom ${userName} bugun qaysi tashkilotlarda bo'lasiz ? `);
    });
    console.log('Напоминания отправлены пользователям');
  } catch (error) {
    console.error('Ошибка при отправке напоминаний:', error);
  }
});
// // Планировщик для отправки напоминания пользователям каждый день в 8:40 по Ташкенту
// const reminderJob = schedule.scheduleJob({ hour: 4, minute:0 , tz: 'Etc/UTC' }, async () => {
//   try {
//     const users = await User.find();
//     users.forEach(user => {
//       bot.sendMessage(user.telegramId, `Assalomu aleykom bugun qaysi tashkilotlarda bo'lasiz ?`);
//     });
//     console.log('Напоминания отправлены пользователям');
//   } catch (error) {
//     console.error('Ошибка при отправке напоминаний:', error);
//   }
// });

// Обработка ошибок
process.on('unhandledRejection', error => {
  console.error('Unhandled promise rejection:', error);
});

bot.on('polling_error', error => {
  console.error('Polling error:', error);
});

console.log(`Bot is starting ...`);
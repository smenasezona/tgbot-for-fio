const {Telegraf} = require('telegraf')
const ExcelJs = require('exceljs')
const {format} = require('date-fns')

const bot = new Telegraf(`6799631932:AAGvoFsC2kptdOMOO1UEsMbGkf6asby9i84`)
const workbook = new ExcelJs.Workbook()

let responseCount = 0
let responseTime = null
const responseLimit = 30

const groupID = -1001848824043

async function searchFioInExcel(name) {
    await workbook.xlsx.readFile("C:\\Users\\splat\\Downloads\\exp2.xlsx")
    const worksheet = workbook.getWorksheet('Лист1')

    const foundedUsers = []

    for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber++) {
        const row = worksheet.getRow(rowNumber);
        if (row.getCell('F').text.trim() === name) {
            const formattedDate = format(row.getCell('A').value, 'dd.MM.yyyy');
            const formattedWeek = format(row.getCell('O').value, 'dd.MM.yyyy');
            const userInfo = {
                name: row.getCell('F').text,
                date: formattedDate,
                payment_type: row.getCell('B').text,
                polis_type: row.getCell('C').text,
                company: row.getCell('D').text,
                contract_number: row.getCell('E').text,
                full_sp: row.getCell('G').text,
                manager: row.getCell('H').text,
                sale_type: row.getCell('C').text,
                director: row.getCell('I').text,
                agent: row.getCell('J').text,
                sale: row.getCell('K').text,
                payment: row.getCell('L').text,
                note: row.getCell('M').text,
                week: row.getCell('N').text,
                expires_date: formattedWeek
            }
            foundedUsers.push(userInfo)
        }
    }
    return foundedUsers.length > 0 ? foundedUsers : null
}

function isResponseLimitExceeded() {
    const currentDate = new Date().toLocaleDateString();
    if (!responseTime || responseTime !== currentDate) {
        responseCount = 0;
        responseTime = currentDate;
        return false;
    }
    return responseCount >= responseLimit;
}

async function isUserInGroup(userId) {
    try {
        const response = await fetch(
            `https://api.telegram.org/bot${bot.token}/getChatMember?chat_id=${groupID}&user_id=${userId}`);
        const result = await response.json();
        return result.ok && (result.result.status === 'member' || result.result.status === 'administrator' || result.result.status === 'creator');
    } catch (error) {
        console.error(error);
        return false;
    }
}

bot.on('new_chat_members', (ctx) => {
    // Получаем ID чата
    const chatId = ctx.chat.id;

    // Отправляем ID чата в чат
    console.log(`ID чата: ${chatId}`)
});

bot.command('search', async (ctx) => {

    const userIsInGroup = await isUserInGroup(ctx.from.id);

    if (!userIsInGroup) {
        ctx.reply('Вы не имеете доступа к этой команде.');
        return;
    }

    if (isResponseLimitExceeded()) {
        ctx.reply(`Превышен лимит ответов. Пожалуйста, подождите.`);
        return;
    }

    const userName = ctx.message.text.split(' ').slice(1).join(' ').trim().toUpperCase();
    const foundUsers = await searchFioInExcel(userName)

    if (foundUsers) {
        foundUsers.forEach(userInfo => {
            ctx.reply(`
ФИО: ${userInfo.name}\n
Дата оплаты: ${userInfo.date}\n
Тип оплаты: ${userInfo.payment_type}\n
Полис: ${userInfo.polis_type}\n
Компания: ${userInfo.company}\n 
Номер договора: ${userInfo.contract_number}\n
Полная СП: ${userInfo.full_sp}\n
Менеджер: ${userInfo.manager}\n
Руководитель: ${userInfo.director}\n
Агент: ${userInfo.agent}\n
Скидка: ${userInfo.sale}\n
Платеж: ${userInfo.payment}\n
Примечание: ${userInfo.note}\n
Неделя: ${userInfo.week}\n
Дата окончания договора: ${userInfo.expires_date}`)
        });
        responseCount++
    } else {
        ctx.reply('Пользователь не найден.');
    }
})

bot.launch()
const {Telegraf} = require('telegraf')
const ExcelJs = require('exceljs')
const {format} = require('date-fns')

const bot = new Telegraf(`6799631932:AAGvoFsC2kptdOMOO1UEsMbGkf6asby9i84`)
const workbook = new ExcelJs.Workbook()

async function searchFioInExcel(name) {
    await workbook.xlsx.readFile("C:\\Users\\splat\\Downloads\\example.xlsx")
    const worksheet = workbook.getWorksheet('Лист1')

    const foundedUsers = []

    for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber++) {
        const row = worksheet.getRow(rowNumber);
        if (row.getCell('H').text.trim() === name) {
            const formattedDate = format(row.getCell('A').value, 'dd.MM.yyyy');
            const userInfo = {
                name: row.getCell('H').text,
                date: formattedDate,
                payment_type: row.getCell('B').text,
                sale_type: row.getCell('C').text,
                polis_type: row.getCell('D').text,
                company: row.getCell('F').text,
                contract_number: row.getCell('G').text,
                full_sp: row.getCell('I').text,
                manager: row.getCell('J').text,
                director: row.getCell('K').text,
                agent: row.getCell('L').text,
                sale: row.getCell('M').text,
                payment: row.getCell('N').text,
                note: row.getCell('O').text,
                week: row.getCell('P').text
            }
            foundedUsers.push(userInfo)
        }
    }
    return foundedUsers.length > 0 ? foundedUsers : null
}

bot.command('search', async (ctx) => {
    const userName = ctx.message.text.split(' ').slice(1).join(' ').trim().toUpperCase();
    const foundUsers = await searchFioInExcel(userName)

    if (foundUsers) {
        foundUsers.forEach(userInfo => {
            ctx.reply(`
ФИО: ${userInfo.name}\n
Дата: ${userInfo.date}\n
Тип оплаты: ${userInfo.payment_type}\n
Тип продажи: ${userInfo.sale_type}\n
Тип полиса: ${userInfo.polis_type}\n
Компания: ${userInfo.company}\n
Номер договора: ${userInfo.contract_number}\n
Полная СП: ${userInfo.full_sp}\n
Менеджер: ${userInfo.manager}\n
Руководитель: ${userInfo.director}\n
Агент: ${userInfo.agent}\n
Скидка: ${userInfo.sale}\n
Платеж: ${userInfo.payment}\n
Примечание: ${userInfo.note}\n
Неделя: ${userInfo.week}`)
        });
    } else {
        ctx.reply('Пользователь не найден.');
    }
})

bot.launch()
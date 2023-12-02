const {Telegraf} = require('telegraf')
const ExcelJs = require('exceljs')

const bot = new Telegraf(`6799631932:AAGvoFsC2kptdOMOO1UEsMbGkf6asby9i84`)
const workbook = new ExcelJs.Workbook()

async function searchFioInExcel(name) {
    await workbook.xlsx.readFile("C:\\Users\\splat\\Downloads\\example.xlsx")
    const worksheet = workbook.getWorksheet('Лист1')
    console.log(workbook.worksheets.map(sheet => sheet.name));

    for (let row of worksheet.getRows(1, 16)) {
        if (row.getCell('H').text === name) {
            return {
                name: row.getCell('H').text,
                date: row.getCell('A').text,
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
        }
    }
    return null
}

bot.command('search', async (ctx) => {
    const userName = ctx.message.text.split(' ')[1]
    const userInfo = await searchFioInExcel(userName)

    if (userInfo) {
        ctx.reply(`ФИО: ${userInfo.name}\n
                   Дата: ${userInfo.date}\n
                   Тип оплаты: ${userInfo.payment_type}\n
                   Тип продажи: ${userInfo.sale_type}\n
                   Тип полиса: ${userInfo.polis_type}\n
                   Компания: ${userInfo.company}\n
                   Номер договора: ${userInfo.contract_number}\n
                   Полная СП: ${userInfo.full_sp}\n
                   Менеджер: ${userInfo.manager}\n
                   Руководитель: ${userInfo.director}\n
                   Агент:${userInfo.agent}\n
                   Скидка: ${userInfo.sale}\n
                   Платеж: ${userInfo.payment}\n
                   Примечание: ${userInfo.note}\n
                   Неделя: ${userInfo.week}`)
    } else {
        ctx.reply('Пользователь не найден.');
    }
})

bot.launch()
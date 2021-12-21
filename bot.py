import openpyxl
import datetime
import telebot
import config

#5063635521:AAH5k6SR3qUHxDs_locfqwhQzhrpGxTsA7g

bot = telebot.TeleBot(config.token)
wb = openpyxl.load_workbook(config.path)

wb.active = 3
ws = wb.active

current_time = '0800'  # str(datetime.datetime.now().strftime('%H%M'))  # текущее время
week_day = datetime.datetime.today().isoweekday()-1  # день недели цифрой

start = [8, 52, 96, 140, 184, 228]  # номер ячейки с которой начинается отсчет в зависимости от дня недели

if week_day == 6:
    time = ['0745', '0845', '0935', '1025', '1115', '1205', '1255']  # расписание субботы

else:
    time = ['0745', '0845', '0935', '1035', '1125', '1215', '1305', '1355', '1440', '1535', '1620', '1705', '1750',
            '1835']  # расписание для всех дней, кроме субботы


@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.send_message(message.chat.id, "Приветсвую!\nОтправьте мне только фамилию учителя и я напишу вам в каком кабинете в данный момент он находится.".format(message.from_user, bot.get_me()), parse_mode='html')

@bot.message_handler(func=lambda m:True)
def search_teacher(message):
    check = 0
    range_cells = 0
    teacher = message.text.lower()

    for i in range(len(time)):
        if time[i] <= current_time < time[i + 1]:
            range_cells = (int(i) + 1) * 3 + start[week_day]

    for row in ws['EV' + str(range_cells):'NJ' + str(range_cells)]:
        for col in row:
            if str(col.value).lower()[:-5] == teacher:
                bot.send_message(message.chat.id, col.offset(row=-1).value)
                check = 1
    if check != 1:
        bot.send_message(message.chat.id, 'Извините, по вашему запросу ничего не найдено🔍\nВозможно, вы неверно ввели фамилию учителя или в данный момент у него нет уроков☹')


if __name__ == '__main__':
    bot.polling(none_stop=True, interval=0)

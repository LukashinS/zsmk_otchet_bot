from openpyxl import load_workbook
import telebot
from datetime import datetime
import json
import os

path = 'bot'


def open_excel_xlsx(path, sheet_number=0):
    """Открывает документ excel (.xlsx) для чтения

    :param path: путь к документу
    :param sheet_number: лист документа
    """

    wb = load_workbook(filename=path)
    sheet = wb.worksheets[sheet_number]

    return wb, sheet


def read_json_from_file(file_name, path=path):
    """Считывание данных из файла
    """

    file_path = os.path.join(path, file_name)
    with open(file_path, encoding='utf-8') as json_file:
        data = json.load(json_file)

    return json.loads(json.dumps(data, ensure_ascii=False))


def write_json(file_name, data_json, path=path):
    """Запись данных в файл
    """

    file_path = os.path.join(path, file_name)
    with open(file_path, 'w', encoding='utf-8') as json_file:
        json.dump(data_json, json_file)


config_dict = read_json_from_file('config.json')
TOKEN_API = config_dict.get('token')
excel_file = os.path.join(path, config_dict.get('excel_file'))
users = config_dict.get('users')
fix_count = ('ввод', 'доп', "снятие", "обсл.", "чистый ввод",
             "инет", "инет+1тв", "новосел", "после новосел", "таунхаус",
             "домофон", "видеодомофон", "etth+1тв", "Etth",
             "доставка сим", "настройка смарт", "85+", "spl1*4", "spl1*8")
edit_count = ("камера", "камера подвес", "кабель", "адваки",
              "шосы", "UTP", "FTP")
area = ("компас", "зсмк", "искитим", "тогучин", "мошково", "коченево", "колывань", "верх-ирмень", "бердск")

bot = telebot.TeleBot(TOKEN_API)


@bot.message_handler(commands=['start'])
def handel_start(message):
    _id = message.from_user.id
    file_name = f"{str(_id)}.json"
    nick = message.from_user.username
    user_name = users.get(nick)
    write_json(file_name, {"Пользователь": user_name if user_name else nick})
    user_markup = telebot.types.ReplyKeyboardMarkup(True, True)
    user_markup.row('Лицевой счет')
    bot.send_message(_id, "Новая заявка", reply_markup=user_markup)


@bot.message_handler(commands=['show'])
def handel_show(message):
    _id = message.from_user.id
    file_name = f"{str(_id)}.json"
    item = read_json_from_file(file_name)
    show_list = [f"{key} - {value}" for key, value in item.items() if
                 not (key in ('id', 'Пользователь', "Переменная")) and not (value == 0)]
    bot.send_message(_id, "\n".join(show_list))


@bot.message_handler(commands=['reset'])
def handel_reset(message):
    _id = message.from_user.id
    user_markup = telebot.types.ReplyKeyboardMarkup(True, True)
    user_markup.row('Лицевой счет')
    bot.send_message(_id, "Новая заявка", reply_markup=user_markup)


@bot.message_handler(commands=['doc'])
def handel_doc(message):
    _id = str(message.from_user.id)
    if _id in config_dict.get("admin"):
        doc = open(excel_file, 'rb')
        bot.send_chat_action(_id, 'upload_document')
        bot.send_document(_id, doc)
        doc.close()
    else:
        bot.send_message(_id, "Не достаточно прав для совершения этого действия")


@bot.message_handler(commands=['finish'])
def handel_finish(message):
    _id = message.from_user.id
    file_name = f"{str(_id)}.json"
    item = read_json_from_file(file_name)

    rec_list = [item.get('Дата'), item.get('Пользователь'), item.get('Лицевой счет'), item.get('Территория'),
                item.get('ввод'), item.get('доп'), item.get('снятие'), item.get('обсл.'), item.get('камера'),
                item.get('камера подвес'), item.get('чистый ввод'), item.get('инет'), item.get('инет+1тв'),
                item.get('85+'), item.get('новосел'), item.get('после новосел'), item.get('таунхаус'),
                item.get('настройка смарт'), item.get('домофон'), item.get('видеодомофон'), item.get('etth+1тв'),
                item.get('Etth'), item.get('доставка сим'), item.get('кабель'), item.get('адваки'),
                item.get('шосы'), item.get('UTP'), item.get('FTP'), item.get('spl1*4'), item.get('spl1*8')]
    try:
        wb, sheet = open_excel_xlsx(excel_file)
        sheet.append([elem if elem and elem != 0 else "" for elem in rec_list])
        wb.save(excel_file)
        user_markup = telebot.types.ReplyKeyboardMarkup(True, True)
        user_markup.row('Лицевой счет')
        bot.send_message(_id, "Новая заявка", reply_markup=user_markup)
    except Exception as inst:
        print('Какая то ошибка: ', inst)
        bot.send_message(_id, "Файл занят, попробуйте добавить позже")


@bot.message_handler(content_types=['text'])
def handel_text(message):
    _id = message.from_user.id

    file_name = f"{str(_id)}.json"
    if not (file_name in os.listdir(path)):
        write_json(file_name, {})
    item = read_json_from_file(file_name)
    msg = message.text

    if item:
        if 'Лицевой счет' in item.keys() and not item.get('Лицевой счет'):
            item['Лицевой счет'] = msg
            item.update(dict.fromkeys(edit_count, 0))
            item.update(dict.fromkeys(fix_count, 0))
            item.update({'Дата': datetime.now().strftime("%d.%m.%Y")})
            item.update({'Территория': None})
            item.update({'Переменная': "Территория"})
            write_json(file_name, item)
            user_markup = telebot.types.ReplyKeyboardMarkup(True)
            user_markup.row(*area[:3])
            user_markup.row(*area[3:6])
            user_markup.row(*area[6:])
            bot.send_message(_id, f"Выбери территорию",
                             reply_markup=user_markup)

    if msg == "Лицевой счет":
        nick = message.from_user.username
        user_name = users.get(nick)
        write_json(file_name, {'Лицевой счет': None, "Пользователь": user_name if user_name else nick})

    if msg in fix_count:
        item[msg] = 1
        write_json(file_name, item)
        bot.send_message(_id, f"Добавил - {msg}")

    variable = item.get('Переменная')
    if variable == "Территория" and (msg in area):
        item["Территория"] = msg
        item['Переменная'] = None
        write_json(file_name, item)
        bot.send_message(_id, f"Добавил '{variable}' - {msg}")
        user_markup = telebot.types.ReplyKeyboardMarkup(True)
        user_markup.row('ввод', 'доп', "снятие", "обсл.")
        user_markup.row("камера", "камера подвес", "чистый ввод")
        user_markup.row("инет", "инет+1тв", "85+", "настройка смарт")
        user_markup.row("новосел", "после новосел", "таунхаус")
        user_markup.row("домофон", "видеодомофон", "etth+1тв", "Etth")
        user_markup.row("доставка сим", "кабель", "адваки")
        user_markup.row("шосы", "UTP", "FTP", "spl1*4", "spl1*8")
        user_markup.row("/show", "/reset", "/finish")
        bot.send_message(_id, f"ЛС:{msg} - Что подключал?",
                         reply_markup=user_markup)

    if variable in edit_count:
        try:
            item[variable] = int(msg)
            item['Переменная'] = None
            write_json(file_name, item)
            bot.send_message(_id, f"Добавил '{variable}' - {msg}")
        except:
            bot.send_message(_id, f"Введите только число")

    if msg in edit_count:
        item['Переменная'] = msg
        write_json(file_name, item)
        bot.send_message(_id, f"Введи кол-во {msg}: ")


if __name__ == "__main__":
    try:
        bot.polling(none_stop=True)
    except Exception as inst:
        print('Какая то ошибка: ', inst)

from openpyxl import load_workbook
import telebot
from datetime import datetime
import json
import os
import sqlite3


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


config_dict = read_json_from_file('config_db.json')
TOKEN_API = config_dict.get('token')
excel_file = os.path.join(path, config_dict.get('excel_file'))
users = config_dict.get('users')
fix_count = ('ввод', 'доп', "снятие", "обсл.", "чистый ввод",
             "инет", "инет+1тв", "новосел", "после новосел", "таунхаус",
             "домофон", "видеодомофон", "etth+1тв", "Etth",
             "доставка сим", "настройка смарт", "85+", "spl1*4", "spl1*8")
edit_count = ("камера", "камера подвес", "кабель", "адваки",
              "шосы", "UTP", "FTP", "ютп", "фтп")
area = ("компас", "зсмк", "искитим", "тогучин", "мошково", "коченево", "колывань", "верх-ирмень", "бердск")

bot = telebot.TeleBot(TOKEN_API)
db_file = os.path.join(path, 'result.db')
db_conn = sqlite3.connect(db_file, check_same_thread=False)
cursor = db_conn.cursor()
# cursor.execute('drop table if exists Tasks')
db_dict = {"id": "id", 'Дата': "date", 'Пользователь': "fio", 'Лицевой счет': "ls", "Территория": "area",
           'ввод': "vvod", 'доп': "dop", "снятие": "snatie", "обсл.": "obsl", "камера": "camera",
           "камера подвес": "camera_podeves", "чистый ввод": "clear_vvod", "инет": "inet", "инет+1тв": "inet_plus_tv",
           "85+": "tv_plus", "новосел": "novosel", "после новосел": "after_novosel", "таунхаус": "townhouse",
           "настройка смарт": "smart", "домофон": "domofon", "видеодомофон": "videodomofon",
           "etth+1тв": "etth_plus_tv", "Etth": "etth", "доставка сим": "sim", "кабель": "cabel", "адваки": "advaki",
           "шосы": "shos", "UTP": "utp", "FTP": "ftp", "spl1*4": "spl14", "spl1*8": "spl18", "Переменная": "variable"}
db_list = db_dict.values()
# cursor.execute(f'CREATE TABLE Tasks {tuple(bd_list)}')


@bot.message_handler(commands=['start'])
def handel_start(message):
    _id = message.from_user.id
    result = cursor.execute(f"SELECT id FROM Tasks WHERE id='{_id}'").fetchone()
    if result:
        cursor.execute(f'DELETE FROM Tasks WHERE "id"={_id}')
    cursor.execute(f'INSERT INTO Tasks (id) VALUES (?)', (_id, ))
    user_markup = telebot.types.ReplyKeyboardMarkup(True, True)
    user_markup.row('Лицевой счет')
    bot.send_message(_id, "Новая заявка", reply_markup=user_markup)


@bot.message_handler(commands=['show'])
def handel_show(message):
    _id = message.from_user.id
    result = cursor.execute(f"SELECT * FROM Tasks WHERE id='{_id}'").fetchone()
    bot.send_message(_id, get_show_result(result))


def get_show_result(result):
    item = {key: result[index] for index, key in enumerate(db_dict)}
    show_list = [f"{key} - {value}" for key, value in item.items() if
                 not (key in ('id', 'Пользователь', "Переменная")) and value]
    return "\n".join(show_list)


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

    db_conn.commit()
    result = cursor.execute(f"SELECT * FROM Tasks WHERE id='{_id}'").fetchone()
    list_res = list(result)
    list_res.pop(0)
    list_res.pop(-1)

    try:
        wb, sheet = open_excel_xlsx(excel_file)
        sheet.append([elem if elem and elem != 0 else "" for elem in list_res])
        wb.save(excel_file)

        print(get_show_result(result))
        print("----------------------")

        user_markup = telebot.types.ReplyKeyboardMarkup(True, True)
        user_markup.row('Лицевой счет')
        bot.send_message(_id, "Новая заявка", reply_markup=user_markup)
    except:
        bot.send_message(_id, "Файл занят, попробуйте добавить позже")


@bot.message_handler(content_types=['text'])
def handel_text(message):
    _id = message.from_user.id
    msg = message.text
    variable = cursor.execute(f"SELECT variable FROM Tasks WHERE id='{_id}'").fetchone()
    if variable and variable[0] == 'Лицевой счет':
        cursor.execute(f'UPDATE Tasks SET {db_dict.get(variable[0])}="{msg}" WHERE "id"={_id}')
        cursor.execute(f'UPDATE Tasks SET variable="Территория" WHERE "id"={_id}')
        bot.send_message(_id, f"Добавил '{variable[0]}' - {msg}")

        user_markup = telebot.types.ReplyKeyboardMarkup(True)
        user_markup.row(*area[:3])
        user_markup.row(*area[3:6])
        user_markup.row(*area[6:])
        bot.send_message(_id, f"Выбери территорию",
                         reply_markup=user_markup)

    if msg == "Лицевой счет":
        nick = message.from_user.username
        user_name = users.get(nick)

        result = cursor.execute(f"SELECT id FROM Tasks WHERE id='{_id}'").fetchone()
        if result:
            cursor.execute(f'DELETE FROM Tasks WHERE "id"={_id}')
        cursor.execute(f'INSERT INTO Tasks (id) VALUES (?)', (_id,))

        cursor.execute(f'UPDATE Tasks SET {db_dict.get("Пользователь")}="{user_name if user_name else nick}", variable="{msg}", date="{datetime.now().strftime("%d.%m.%Y")}" WHERE "id"={_id}')

    if variable and variable[0] == "Территория" and (msg in area):
        cursor.execute(f'UPDATE Tasks SET area="{msg}" WHERE "id"={_id}')
        cursor.execute(f'UPDATE Tasks SET variable="" WHERE "id"={_id}')
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

    if msg in fix_count:
        cursor.execute(f'UPDATE Tasks SET {db_dict.get(msg)}="1" WHERE "id"={_id}')
        bot.send_message(_id, f"Добавил - {msg}")

    if variable and variable[0] in edit_count:
        var = variable[0]
        if var == 'ютп':
            value = 'UTP'
        elif var == 'фтп':
            value = 'FTP'
        else:
            value = var
        try:
            cursor.execute(f'UPDATE Tasks SET {db_dict.get(value)}="{int(msg)}" WHERE "id"={_id}')
            cursor.execute(f'UPDATE Tasks SET variable="" WHERE "id"={_id}')
            bot.send_message(_id, f"Добавил '{value}' - {msg}")
        except:
            bot.send_message(_id, f"Введите только число")

    if msg in edit_count:
        if msg == 'UTP':
            value = 'ютп'
        elif msg == 'FTP':
            value = 'фтп'
        else:
            value = msg
        cursor.execute(f'UPDATE Tasks SET variable="{value}" WHERE "id"={_id}')
        bot.send_message(_id, f"Введи кол-во {msg}: ")


if __name__ == "__main__":
    try:
        bot.polling(none_stop=True)
    except Exception as inst:
        print('Какая то ошибка: ', inst)

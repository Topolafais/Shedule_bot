import os
import json
import telebot
from telebot import types
from datetime import datetime, timedelta
import openpyxl
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
from pydrive2.settings import LoadSettingsFile
from dotenv import load_dotenv
from oauth2client.service_account import ServiceAccountCredentials

load_dotenv()

if os.getenv("RENDER"):
    service_account_info = json.loads(os.getenv("GOOGLE_SERVICE_ACCOUNT"))
else:
    with open("service_account.json", "r", encoding="utf-8") as f:
        service_account_info = json.load(f)

creds = ServiceAccountCredentials.from_json_keyfile_dict(
    service_account_info, scopes=["https://www.googleapis.com/auth/drive"]
)

gauth = GoogleAuth()
gauth.LoadCredentialsFile = None
gauth.settings = LoadSettingsFile()
gauth.ServiceAuth(service_account_info)

drive = GoogleDrive(gauth)

token = os.getenv("token")

month_names = {
    1: "січень", 2: "лютий", 3: "березень", 4: "квітень",
    5: "травень", 6: "червень", 7: "липень", 8: "серпень",
    9: "вересень", 10: "жовтень", 11: "листопад", 12: "грудень"
}

folder = "1DvZAowwZRTFtCnwSUDK_9PHtD71u__-R"
t = ["\n  --== 8:30 ==--\n", "\n  --== 10:05 ==--\n", "\n  --== 11:55 ==--\n", "\n  --== 13:30 ==--\n", "\n  --== 15:05 ==--\n"]
tomorrow = datetime.now() + timedelta(days=1)
month_name = month_names[tomorrow.month]
date_str = tomorrow.strftime("%d")
stolb = 6


def find_file():
    folders = drive.ListFile({'q': f"'{folder}' in parents and trashed=false and mimeType='application/vnd.google-apps.folder'"}).GetList()

    month_folder = None
    for f in folders:
        if f['title'].lower() == month_name.lower():
            month_folder = f
            break

    if not month_folder:
        raise Exception(f"Папку '{month_name}' не знайдено!")

    print(f"Знайдено папку місяця: {month_folder['title']}")

    files = drive.ListFile({'q': f"'{month_folder['id']}' in parents and trashed=false"}).GetList()

    today_file = None
    for f in files:
        if f['title'].startswith(date_str) and f['title'].endswith('.xlsx'):
            today_file = f
            break

    if not today_file:
        raise Exception(f"Файл на {date_str} число не знайдено!")

    print(f"Знайдено файл: {today_file['title']}")

    local_path = os.path.join(os.getcwd(), "temp.xlsx")
    today_file.GetContentFile(local_path)
    file = openpyxl.load_workbook(local_path, read_only=True)
    sheet = file.active
    return sheet


bot = telebot.TeleBot(token)
@bot.message_handler(commands=['start'])
def start(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    item1 = types.KeyboardButton("Розклад на завтра")
    item2 = types.KeyboardButton("Вказати группу")
    markup.add(item1)
    markup.add(item2)
    bot.send_message(message.chat.id, "Привіт. За допомогою цього бота ти можешь дізнатися розклад на завтра!", reply_markup=markup)

@bot.message_handler(content_types=['text'])
def reply(message):
    global stolb
    if message.text == "Розклад на завтра":
        sheet = find_file()
        arr = []
        for i in range(16, 25, 2):
            if stolb // 2 != 2:
                if sheet[i][stolb] in sheet.merged_cells:
                    arr.append(t[int((i - 14) / 2) - 1])
                    arr.append(sheet[i][stolb - 1].value)
            elif sheet[i][stolb].value is not None:
                arr.append(t[int((i - 14) / 2) - 1])
                arr.append(sheet[i][stolb].value)
        bot.send_message(message.chat.id, '\n'.join(map(str, arr)))
    elif message.text == "Вказати группу":
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        item1 = types.KeyboardButton("Ic 2/1")
        item2 = types.KeyboardButton("Ic 2/2")
        markup.add(item1)
        markup.add(item2)
        bot.send_message(message.chat.id, "Вибери свою группу", reply_markup=markup)
    elif message.text == "Ic 2/1":
        stolb = 6
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        item1 = types.KeyboardButton("Розклад на завтра")
        item2 = types.KeyboardButton("Вказати группу")
        markup.add(item1)
        markup.add(item2)
        bot.send_message(message.chat.id, "Зрозумів, ти з Ic 2/1", reply_markup=markup)
    elif message.text == "Ic 2/2":
        stolb = 7
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        item1 = types.KeyboardButton("Розклад на завтра")
        item2 = types.KeyboardButton("Вказати группу")
        markup.add(item1)
        markup.add(item2)
        bot.send_message(message.chat.id, "Зрозумів, ти з Ic 2/2", reply_markup=markup)


bot.infinity_polling()
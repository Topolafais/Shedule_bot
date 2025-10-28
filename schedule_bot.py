import os

import telebot
from telebot import types
from datetime import datetime, timedelta
import openpyxl
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import pandas as pd
from io import BytesIO

month_names = {
    1: "січень", 2: "лютий", 3: "березень", 4: "квітень",
    5: "травень", 6: "червень", 7: "липень", 8: "серпень",
    9: "вересень", 10: "жовтень", 11: "листопад", 12: "грудень"
}

gauth = GoogleAuth()
gauth.LocalWebserverAuth()
drive = GoogleDrive(gauth)
token = "8307327977:AAFx1PCWYzDJ6dvQJHZUik4kh8ZaUAuLqp4"
folder = "1DvZAowwZRTFtCnwSUDK_9PHtD71u__-R"
t = ["8:30", "10:05", "11:55", "13:30", "15:05"]
tomorrow = datetime.now() + timedelta(days=1)
month_name = month_names[tomorrow.month]
date_str = tomorrow.strftime("%d")
global stolb
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
    #item2 = types.KeyboardButton("Вказати группу")
    markup.add(item1)
    #markup.add(item2)
    bot.send_message(message.chat.id, "Привіт. За допомогою цього бота ти можешь дізнатися розклад на завтра!", reply_markup=markup)

@bot.message_handler(content_types=['text'])
def reply(message):
    if message.text == "Розклад на завтра":
        sheet = find_file()
        arr = []
        for i in range(16, 25, 2):
            if sheet[i][stolb].value is not None:
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

@bot.message_handler(content_types=['text'])
def group(message):
    if message.text == "Ic 2/1":
        stolb = 6
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        item1 = types.KeyboardButton("Розклад на завтра")
        item2 = types.KeyboardButton("Вказати группу")
        markup.add(item1)
        markup.add(item2)
    elif message.text == "Ic 2/1":
        stolb = 7
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        item1 = types.KeyboardButton("Розклад на завтра")
        item2 = types.KeyboardButton("Вказати группу")
        markup.add(item1)
        markup.add(item2)


bot.infinity_polling()
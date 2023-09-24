import telebot
import os
import docx
from telebot import types
import pandas as pd
from prettytable import PrettyTable
import io  # Добавьте импорт модуля io


# Your Telegram bot token
TOKEN = ''

# Initialize the bot
bot = telebot.TeleBot(TOKEN)
@bot.message_handler(commands=['start'])
def start(message):
    chat_id= message.chat.id
    markup = types.InlineKeyboardMarkup()
    word_button = types.InlineKeyboardButton("Word", callback_data='word')
    pdf_button = types.InlineKeyboardButton("excel", callback_data='excel')
    markup.add(word_button, pdf_button)
    bot.send_message(chat_id, "Выберите формат файла:", reply_markup=markup)


# Handle documents sent by users
def handle_word_document(message):
    try:
        # Download the document
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)

        # Save the document locally
        with open(message.document.file_name, 'wb') as doc_file:
            doc_file.write(downloaded_file)

        # Extract text from the document
        doc = docx.Document(message.document.file_name)
        text = "\n".join([paragraph.text for paragraph in doc.paragraphs])

        # Send the extracted text back to the user
        bot.send_message(message.chat.id, f'Text extracted from "{message.document.file_name}":\n{text}')

        # Clean up by deleting the local file
        os.remove(message.document.file_name)
    except Exception as e:
        print(str(e))
        bot.send_message(message.chat.id, 'An error occurred while processing the document.')


def handle_document(message):
    try:
        chat_id = message.chat.id

        # Получаем информацию о файле
        file_info = bot.get_file(message.document.file_id)
        file_path = file_info.file_path

        # Скачиваем файл
        downloaded_file = bot.download_file(file_path)

        # Читаем содержимое Excel-файла с использованием pandas и io.BytesIO
        with io.BytesIO(downloaded_file) as bio:
            df = pd.read_excel(bio)

        # Преобразуем DataFrame в красивую таблицу с помощью PrettyTable
        table = PrettyTable(header_style='title')
        table.field_names = df.columns
        for row in df.itertuples(index=False):
            table.add_row(row)

        # Отправляем таблицу пользователю
        bot.send_message(chat_id, table.get_string(), parse_mode='HTML')

    except Exception as e:
        print( f'Произошла ошибка: {str(e)}')

@bot.callback_query_handler(func=lambda call: True)
def handle_callback_query(call):
    if call.data == 'word':
        bot.send_message(call.message.chat.id, "Вы выбрали формат Word. Пожалуйста, отправьте файл Word.")
        bot.register_next_step_handler(call.message,handle_word_document)
    elif call.data == 'excel':
        bot.send_message(call.message.chat.id, "Вы выбрали формат excel. Пожалуйста, отправьте файл PDF.")
        bot.register_next_step_handler(call.message,handle_document)

# Start the bot
if __name__ == '__main__':
    bot.polling(none_stop=True)

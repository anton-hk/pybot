import telebot
from telebot import types
import pythoncom
import os, shutil
from Conversation import Convert
bot = telebot.TeleBot("your tokin")

# Приветствие
@bot.message_handler(commands=['start'])
def send_welcome(message):
    sti = open('sticker.webp', 'rb')
    bot.send_sticker(message.chat.id, sti)
    bot.send_message(message.chat.id, "Hello, {0.first_name}!\nI'm — <b>{1.first_name}</b>,"
                                      "a bot that converts photos and docx, doc, docm, jpg, png files to pdf"
                                      "".format(message.from_user, bot.get_me()), parse_mode='html')
    sti.close()


@bot.message_handler(content_types=['text'])
def get_text_messages(message):
    bot.send_message(message.chat.id, "Send me photo or file")



@bot.message_handler(content_types=['document'])
def get_text_messages(message):
    try:
        conv = Convert()
        doc_format = conv.parse_name(message.document.file_name)[1]  # формат дакумента (docx) (без точки)
        name = conv.parse_name(message.document.file_name)[0]  # получаем строку с названием файла без .docx
        file_info = bot.get_file(message.document.file_id)
        download_file = bot.download_file(file_info.file_path)
        with open(r'C:\Users\Lumpen\PycharmProjects\pdf_bot\%(name)s'
                  % {"name": str(message.document.file_name)}, 'wb') as new_file:
            new_file.write(download_file)  # качаем файл который скинули
        pythoncom.CoInitializeEx(0)
        if doc_format == "docx" or doc_format == "doc" or doc_format == "docm":
            conv.conversation_word(str(message.document.file_name))  # конвертируем вордовский документ
        if doc_format == 'jpg' or doc_format == 'png':
            conv.conversation_jpg(str(message.document.file_name))  # конвертируем файл jpg
        doc = open(r'C:\Users\Lumpen\PycharmProjects\pdf_bot\%(name)s.pdf' % {"name": name}, 'rb')
        bot.send_document(message.chat.id, doc)  # отправляем
        doc.close()
        new_file.close()
        os.remove(r'C:\Users\Lumpen\PycharmProjects\pdf_bot\%(name)s.pdf' % {"name": name})
        os.remove(r'C:\Users\Lumpen\PycharmProjects\pdf_bot\%(name)s' % {"name": message.document.file_name})
    except:
        bot.send_message(message.chat.id, "Sorry, I cannot convert this or some error occurred")


@bot.message_handler(content_types=['photo'])
def get_text_messages(message):
    file_info = bot.get_file(message.photo[1].file_id)
    downloaded_file = bot.download_file(file_info.file_path)
    file_path = 'C:/Users/Lumpen/PycharmProjects/pdf_bot/'
    src = file_path + file_info.file_path
    with open(src, 'wb') as new_file:
        new_file.write(downloaded_file) #  Загрузка всех фото в папку photos
    markup = types.InlineKeyboardMarkup(row_width=1)  #  Создание клавиатуры бота
    item1 = types.InlineKeyboardButton('Convert', callback_data='convert')
    markup.add(item1)
    bot.send_message(message.chat.id, 'Done, push the button', reply_markup=markup) # Отправка сообщения с прикрепленной к нему клавиатурой


@bot.callback_query_handler(func=lambda call: True)
def callback_inline(call):
    try:
        if call.message:
            if call.data == 'convert':
                conv = Convert()
                conv.conversation_list_images() #  Конвертация всех фото
                doc = open(r'C:\Users\Lumpen\PycharmProjects\pdf_bot\photos\file.pdf', 'rb')
                bot.send_document(call.message.chat.id, doc)
                doc.close()
        # bot.edit_message_reply_markup(chat_id=call.message.chat.id, message_id=call.message.chat.id, reply_markup=None)
        path = r'C:\Users\Lumpen\PycharmProjects\pdf_bot\photos'
        for filename in os.listdir(path): # Удаление всех файлов в папке photos
            file_path = os.path.join(path, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print('Failed to delete %s. Reason: %s' % (file_path, e))
    except:
        pass

bot.polling(none_stop=True, interval=0)

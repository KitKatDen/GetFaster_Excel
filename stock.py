import telebot
import wikipedia
from telebot import apihelper
apihelper.proxy = {'http':'http://141.164.42.178:8080'}

wikipedia.set_lang("ru")

bot = telebot.TeleBot("token");
@bot.message_handler(content_types=['text'])
def get_text_messages(message):
    if message.text == "test":
        bot.send_message(message.chat.id, wikipedia.summary("ИС-2"))
bot.polling(none_stop=True, interval=0)

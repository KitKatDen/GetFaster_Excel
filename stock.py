import telebot
import requests
import wikipedia
#import telepot
from bs4 import BeautifulSoup

html = requests.get("https://ru.wikipedia.org/wiki/ИС-2")
wikipedia.set_lang("ru")

bot = telebot.TeleBot('1244076324:AAGx55Gq-HMB0y5wdI3uanCErd5BnnvGRYw');
@bot.message_handler(content_types=['text'])
def get_text_messages(message):
    if message.text == "test":
        bot.send_message(message.chat.id, wikipedia.summary("ИС-2"))
bot.polling(none_stop=True, interval=0)

message.from_chat.id
print(bot.get_updates()[0].message.chat.id)

from dotenv import load_dotenv, dotenv_values
import telebot 
import sqlite3

load_dotenv()
config = dotenv_values(".env")

con = sqlite3.connect("datausr.db", check_same_thread=False)


API_TOKEN = config['TOKEN']

bot = telebot.TeleBot(API_TOKEN)

# Handle '/start' and '/help'
@bot.message_handler(commands=['help', 'start'])
def send_welcome(message):
    bot.reply_to(message, """\
Hi there, I am EchoBot.
I am here to echo your kind words back to you. Just say anything nice and I'll say the exact same thing to you!\
""")
    
@bot.message_handler(commands=['get'])
def get_worker(message):
    cur = con.cursor()
    s = message.text.split(' ', 1)
    if len(s) < 2:
        bot.reply_to(message, 'Введите ФИО')
        return
    res = cur.execute(f"select * from datauser where FIO = '{s[1]}'")
    data = res.fetchone()
    if data is None:
        bot.reply_to(message, 'Рабочий не найден')
        return
    bot.reply_to(message, str(data))

    
# Handle all other messages with content_type 'text' (content_types defaults to ['text'])
@bot.message_handler(func=lambda message: True)
def echo_message(message):
    bot.reply_to(message, message.text)


bot.infinity_polling()

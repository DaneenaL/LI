from datetime import datetime
from docx import Document
from dotenv import load_dotenv, dotenv_values
import telebot 
import sqlite3
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt


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
    doc = Document("LI Pochta+Eosdo+1c.docx")
    styles = doc.styles
    BIGstyle = styles.add_style('Big', WD_STYLE_TYPE.PARAGRAPH)
    BIGstyle.font.name = 'Times New Roman'
    BIGstyle.font.size = Pt(12)
    SMALLstyle = styles.add_style('Small', WD_STYLE_TYPE.PARAGRAPH)
    SMALLstyle.font.name = 'Times New Roman'
    SMALLstyle.font.size = Pt(11)

    cur = con.cursor()
    s = message.text.split(' ', 1)
    if len(s) < 2:
        bot.reply_to(message, 'Введите ФИО')
        return
    res = cur.execute(f"select * from datauser where FIO = '{s[1]}'")
    data = res.fetchone()
    if data is None:
        bot.reply_to(message, "Работник не найден")
        return
    FIO = data[1]
    number = data[2]
    position = data[3]
    department = data[4]
    address = data[5]
    director = data[6]
    date = datetime.today().strftime("%d.%m.%Y")
    name = FIO.split(" ") 
    name = name[0] + " " + name[1][0] + "." + (name[2][0] + "." if len(name) > 2 else "")
    doc.tables[0].rows[1].cells[1].paragraphs[1].text = FIO
    doc.tables[0].rows[1].cells[1].paragraphs[1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.tables[0].rows[1].cells[1].paragraphs[1].style = BIGstyle
    doc.tables[0].rows[5].cells[1].paragraphs[0] = str(number) 
    doc.tables[0].rows[5].cells[1].paragraphs[1].text.style = "\nУкажите обязательно для учетной записи категории А, для других категорий оставьте поле пустым"       
    # doc.tables[0].rows[6].cells[1].paragraphs = f"\n{position}\nУкажите обязательно для учетной записи категории А, для других категорий оставьте поле пустым"
    # doc.tables[0].rows[8].cells[1].paragraphs = f'{department}\nДля категории А – в соответствии со штатным расписанием, для остальных – куда прикрепляется'
    # doc.tables[0].rows[9].cells[1].paragraphs = f'{address}\nУкажите фактический адрес рабочего места пользователя и номер кабинета'
    # doc.tables[0].rows[14].cells[1].paragraphs = f'\n{name}\nРасшифровка подписи (ФИО)'
    # doc.tables[0].rows[14].cells[4].paragraphs = f'\n{date}\nДата подписания'

    # doc.tables[2].rows[8].cells[2].paragraphs = f'\n{director}\nФИО'
    # doc.tables[2].rows[8].cells[4].paragraphs = f'\n{date}\nДата подписания'
    # doc.tables[2].rows[9].cells[4].paragraphs = f'\n{date}\nДата подписания'
    # doc.tables[2].rows[17].cells[1].paragraphs = f'\n{name}\nРасшифровка подписи (ФИО)'
    # doc.tables[2].rows[17].cells[3].paragraphs = f'\n{date}\nДата подписания'
    # doc.tables[2].rows[3].cells[4].paragraphs = f'\n{date}\nДата подписания'

    # doc.tables[4].rows[8].cells[4].paragraphs = f'\n{date}\nДата подписания'
    # doc.tables[4].rows[9].cells[4].paragraphs = f'\n{date}\nДата подписания'
    # doc.tables[4].rows[8].cells[2].paragraphs = f'\n{director}\nФИО'
    # doc.tables[4].rows[17].cells[1].paragraphs = f'\n{name}\nРасшифровка подписи (ФИО)'
    # doc.tables[4].rows[17].cells[3].paragraphs = f'\n{date}\nДата подписания'
    # doc.tables[4].rows[3].cells[4].paragraphs = f'\n{date}\nДата подписания'

    # doc.tables[6].rows[8].cells[2].paragraphs = f'\n{director}\nФИО'
    # doc.tables[6].rows[8].cells[4].paragraphs = f'\n{date}\nДата подписания'
    # doc.tables[6].rows[9].cells[4].paragraphs = f'\n{date}\nДата подписания'
    # doc.tables[6].rows[17].cells[1].paragraphs = f'\n{name}\nРасшифровка подписи (ФИО)'
    # doc.tables[6].rows[17].cells[3].paragraphs = f'\n{date}\nДата подписания'
    # doc.tables[6].rows[3].cells[4].paragraphs = f'\n{date}\nДата подписания'

    doc.save(f"LI_{name}.doc")

# Handle all other messages with content_type 'text' (content_types defaults to ['text'])
@bot.message_handler(func=lambda message: True)
def echo_message(message):
    bot.reply_to(message, message.text)


bot.infinity_polling()

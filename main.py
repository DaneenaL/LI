from datetime import datetime
import os
from docx import Document
from dotenv import load_dotenv, dotenv_values
import telebot 
import sqlite3
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt
from docx.text.paragraph import Paragraph
from docx.styles.style import BaseStyle


load_dotenv()
config = dotenv_values(".env")

con = sqlite3.connect("datausr.db", check_same_thread=False)

API_TOKEN = config['TOKEN']

bot = telebot.TeleBot(API_TOKEN)

# Handle '/start' and '/help'
@bot.message_handler(commands=['help', 'start'])
def send_welcome(message):
    bot.reply_to(message, """\
Напишите |/get ФИО| работника, на которого нужен лист исполнения\
""")
    
def apply_style(paragraph: Paragraph, text: str, style: BaseStyle):
    paragraph.text = text
    paragraph.style = style
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

def build_styles(doc):
    styles = doc.styles
    BIGstyle = styles.add_style('Big', WD_STYLE_TYPE.PARAGRAPH)
    BIGstyle.font.name = 'Times New Roman'
    BIGstyle.font.size = Pt(12)
    SMALLstyle = styles.add_style('Small', WD_STYLE_TYPE.PARAGRAPH)
    SMALLstyle.font.name = 'Times New Roman'
    SMALLstyle.font.size = Pt(11)

    return BIGstyle, SMALLstyle
    
@bot.message_handler(commands=['get'])
def get_worker(message):
    doc = Document("LI Pochta+Eosdo+1c.docx")

    BIGstyle, SMALLstyle = build_styles(doc)

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

    apply_style(doc.tables[0].rows[1].cells[1].paragraphs[1], FIO, BIGstyle)
    apply_style(doc.tables[0].rows[5].cells[1].paragraphs[0], str(number), BIGstyle)
    apply_style(doc.tables[0].rows[6].cells[1].paragraphs[1], position, BIGstyle)
    apply_style(doc.tables[0].rows[8].cells[1].paragraphs[0], department, BIGstyle)
    apply_style(doc.tables[0].rows[9].cells[1].paragraphs[1], address, BIGstyle)
    apply_style(doc.tables[0].rows[9].cells[1].paragraphs[0], '', BIGstyle)
    apply_style(doc.tables[0].rows[14].cells[1].paragraphs[1], name, BIGstyle)
    apply_style(doc.tables[0].rows[14].cells[4].paragraphs[1], date, BIGstyle)
    

    apply_style(doc.tables[2].rows[8].cells[2].paragraphs[0], director, BIGstyle)
    apply_style(doc.tables[2].rows[8].cells[4].paragraphs [0], date, BIGstyle)
    apply_style(doc.tables[2].rows[9].cells[4].paragraphs [0], date, BIGstyle)
    apply_style(doc.tables[2].rows[17].cells[1].paragraphs[0], name, BIGstyle)
    apply_style(doc.tables[2].rows[17].cells[3].paragraphs[0], date, BIGstyle)
    apply_style(doc.tables[2].rows[3].cells[4].paragraphs [0], date, BIGstyle)

    apply_style(doc.tables[4].rows[8].cells[2].paragraphs[0], director, BIGstyle)
    apply_style(doc.tables[4].rows[8].cells[4].paragraphs [0], date, BIGstyle)
    apply_style(doc.tables[4].rows[9].cells[4].paragraphs [0], date, BIGstyle)
    apply_style(doc.tables[4].rows[17].cells[1].paragraphs[0], name, BIGstyle)
    apply_style(doc.tables[4].rows[17].cells[3].paragraphs[0], date, BIGstyle)
    apply_style(doc.tables[4].rows[3].cells[4].paragraphs [0], date, BIGstyle)

   

    apply_style(doc.tables[6].rows[8].cells[2].paragraphs[0], director, BIGstyle)
    apply_style(doc.tables[6].rows[8].cells[4].paragraphs [0], date, BIGstyle)
    apply_style(doc.tables[6].rows[9].cells[4].paragraphs [0], date, BIGstyle)
    apply_style(doc.tables[6].rows[17].cells[1].paragraphs[0], name, BIGstyle)
    apply_style(doc.tables[6].rows[17].cells[3].paragraphs[0], date, BIGstyle)
    apply_style(doc.tables[6].rows[3].cells[4].paragraphs [0], date, BIGstyle)

    filename = f"LI_{name}.doc"
    doc.save(filename)
    bot.send_document(message.chat.id, open(filename, 'rb'))
    os.remove(filename)


bot.infinity_polling()
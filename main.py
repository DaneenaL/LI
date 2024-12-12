from datetime import datetime
import os
from docx import Document
from dotenv import load_dotenv, dotenv_values
import telebot 
import psycopg2
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt
from docx.text.paragraph import Paragraph
from docx.styles.style import BaseStyle

from consts import BASE_TEMPLATE_FOLDER

import os.path


load_dotenv()
config = dotenv_values(".env")

API_TOKEN = config['TOKEN']

bot = telebot.TeleBot(API_TOKEN)

# Handle '/start' and '/help'
@bot.message_handler(commands=['help', 'start'])
def send_welcome(message):
    bot.reply_to(message, """\
/gli ФИО - Лист исполнения на почту, 1С, ЕОСДО, BOXER
/glip ФИО - Лист исполнения на почту
/glic ФИО - Лист исполнения на 1С
/glice ФИО - Лист исполнения на 1С, ЕОСДО
/glipc ФИО - Лист исполнения на 1С, Почту
/glie ФИО - Лист исполнения на ЕОСДО
/gliep ФИО - Лист исполнения на ЕОСДО, Почту
/glib ФИО - Лист исполнения на BOXER
/gsz ФИО - Служебка на УЗ в kisozk.local
/glipce ФИО - Лист исполнения на почту, 1С, ЕОСДО
/glicb ФИО - Лист исполнения на 1С, BOXER
/gliceb ФИО - Лист исполнения на 1С, ЕОСДО, BOXER
/glipcb ФИО - Лист исполнения на 1С, Почту, BOXER
/glieb ФИО - Лист исполнения на ЕОСДО, BOXER
/gliepb ФИО - Лист исполнения на ЕОСДО, Почту, BOXER
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

    return BIGstyle

def select_from_datauser(value: str):
    with psycopg2.connect(dbname = config["POSTGRES_DB"], user = config["POSTGRES_USER"], password = config["POSTGRES_PASSWORD"], host = "localhost", port = "5432") as con:
        with con.cursor() as cur:
            cur.execute('select * from workers where username = %s', (value,))
            result = cur.fetchall()
    return result
    
def check_permissions(telegram_ID: int) -> bool:
    with psycopg2.connect(dbname = config["POSTGRES_DB"], user = config["POSTGRES_USER"], password = config["POSTGRES_PASSWORD"], host = "localhost", port = "5432") as con:
        with con.cursor() as cur:
            cur.execute('select * from users where telegram_id = %s', (telegram_ID,))
            result = cur.fetchone()
    return bool(result)



#list ispolneniya na 1C
@bot.message_handler(commands=['glic'])
def get_worker(message):
    if not check_permissions(message.from_user.id):
        bot.reply_to(message, "Доступа нет")
        return
    s = message.text.split(' ', 1)
    if len(s) < 2:
        bot.reply_to(message, 'Введите ФИО')
        return
    
    data = select_from_datauser(s[1])
    if not data:
        bot.reply_to(message, "Работник не найден")
        return
    data = data[0] # TODO: поставить обработку нескольких работников

    fio = data[1]
    number = data[2]
    position = data[3]
    department = data[4]
    address = data[5]
    director = data[6]
    date = datetime.today().strftime("%d.%m.%Y")
    name = fio.split(" ") 
    name = name[0] + " " + name[1][0] + "." + (name[2][0] + "." if len(name) > 2 else "")

    doc = Document(os.path.join(BASE_TEMPLATE_FOLDER, "LI_1C.docx"))
    BIGstyle = build_styles(doc)

    apply_style(doc.tables[0].rows[1].cells[1].paragraphs[1], fio, BIGstyle)
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
    
    filename = f"LI_{name}_1C.doc"
    doc.save(filename)
    bot.send_document(message.chat.id, open(filename, 'rb'))
    os.remove(filename)

    #list ispolneniya na EOSDO
@bot.message_handler(commands=['glie'])
def get_worker(message):
    doc = Document("LI_EOSDO.docx")

    BIGstyle = build_styles(doc)

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
    fio = data[1]
    number = data[2]
    position = data[3]
    department = data[4]
    address = data[5]
    director = data[6]
    date = datetime.today().strftime("%d.%m.%Y")
    name = fio.split(" ") 
    name = name[0] + " " + name[1][0] + "." + (name[2][0] + "." if len(name) > 2 else "")

    apply_style(doc.tables[0].rows[1].cells[1].paragraphs[1], fio, BIGstyle)
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
    
    filename = f"LI_{name}_EOSDO.doc"
    doc.save(filename)
    bot.send_document(message.chat.id, open(filename, 'rb'))
    os.remove(filename)


# list ispolneniya na vse
@bot.message_handler(commands=['gli'])
def get_worker(message):
    doc = Document("LI_Pochta+EOSDO+1C+BOXER.docx")

    BIGstyle = build_styles(doc)

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
    fio = data[1]
    number = data[2]
    position = data[3]
    department = data[4]
    address = data[5]
    director = data[6]
    date = datetime.today().strftime("%d.%m.%Y")
    name = fio.split(" ") 
    name = name[0] + " " + name[1][0] + "." + (name[2][0] + "." if len(name) > 2 else "")

    apply_style(doc.tables[0].rows[1].cells[1].paragraphs[1], fio, BIGstyle)
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

    apply_style(doc.tables[8].rows[8].cells[2].paragraphs[0], director, BIGstyle)
    apply_style(doc.tables[8].rows[8].cells[4].paragraphs [0], date, BIGstyle)
    apply_style(doc.tables[8].rows[9].cells[4].paragraphs [0], date, BIGstyle)
    apply_style(doc.tables[8].rows[17].cells[1].paragraphs[0], name, BIGstyle)
    apply_style(doc.tables[8].rows[17].cells[3].paragraphs[0], date, BIGstyle)
    apply_style(doc.tables[8].rows[3].cells[4].paragraphs [0], date, BIGstyle)
    
    filename = f"LI_{name}_Pochta_Eosdo_1C_BOXER.doc"
    doc.save(filename)
    bot.send_document(message.chat.id, open(filename, 'rb'))
    os.remove(filename)

#sluzhebka na kisozk
@bot.message_handler(commands=['gsz'])
def get_worker_sz_istok(message):
    doc = Document("SZ_istok_na_UZ.docx")

    BIGstyle = build_styles(doc)

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
    fio = data[1]
    number = data[2]
    position = data[3]
    department = data[4]
    name = fio.split(" ") 
    name = name[0] + " " + name[1][0] + "." + (name[2][0] + "." if len(name) > 2 else "")
    
    apply_style(doc.tables[1].rows[1].cells[0].paragraphs[0], fio, BIGstyle)
    apply_style(doc.tables[1].rows[1].cells[1].paragraphs[0], department, BIGstyle)
    apply_style(doc.tables[1].rows[1].cells[2].paragraphs[0], str(number), BIGstyle)
    doc.paragraphs[8].text =  position+' '*(123-len(position)-len(fio))+fio
    doc.paragraphs[8].style = BIGstyle  
    doc.paragraphs[8].alignment = WD_ALIGN_PARAGRAPH.LEFT

    filename = f"LI_{name}.doc"
    doc.save(filename)
    bot.send_document(message.chat.id, open(filename, 'rb'))
    os.remove(filename)


#list ispolneniya na pochtu
@bot.message_handler(commands=['glip'])
def get_worker(message):
    doc = Document("LI_Pochta.docx")

    BIGstyle = build_styles(doc)

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
    fio = data[1]
    number = data[2]
    position = data[3]
    department = data[4]
    address = data[5]
    director = data[6]
    date = datetime.today().strftime("%d.%m.%Y")
    name = fio.split(" ") 
    name = name[0] + " " + name[1][0] + "." + (name[2][0] + "." if len(name) > 2 else "")

    apply_style(doc.tables[0].rows[1].cells[1].paragraphs[1], fio, BIGstyle)
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
    
    filename = f"LI_{name}_Pochta.doc"
    doc.save(filename)
    bot.send_document(message.chat.id, open(filename, 'rb'))
    os.remove(filename)


#list ispolneniya na 1C EOSDO
@bot.message_handler(commands=['glice'])
def get_worker(message):
    doc = Document("LI_EOSDO+1C.docx")

    BIGstyle = build_styles(doc)

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
    fio = data[1]
    number = data[2]
    position = data[3]
    department = data[4]
    address = data[5]
    director = data[6]
    date = datetime.today().strftime("%d.%m.%Y")
    name = fio.split(" ") 
    name = name[0] + " " + name[1][0] + "." + (name[2][0] + "." if len(name) > 2 else "")

    apply_style(doc.tables[0].rows[1].cells[1].paragraphs[1], fio, BIGstyle)
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
    
    filename = f"LI_{name}_1C_EOSDO.doc"
    doc.save(filename)
    bot.send_document(message.chat.id, open(filename, 'rb'))
    os.remove(filename)

#list ispolneniya na 1C+Pochta
@bot.message_handler(commands=['glipc'])
def get_worker(message):
    doc = Document("LI_Pochta+1C.docx")

    BIGstyle = build_styles(doc)

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
    fio = data[1]
    number = data[2]
    position = data[3]
    department = data[4]
    address = data[5]
    director = data[6]
    date = datetime.today().strftime("%d.%m.%Y")
    name = fio.split(" ") 
    name = name[0] + " " + name[1][0] + "." + (name[2][0] + "." if len(name) > 2 else "")

    apply_style(doc.tables[0].rows[1].cells[1].paragraphs[1], fio, BIGstyle)
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
    
    filename = f"LI_{name}_1C_Pochta.doc"
    doc.save(filename)
    bot.send_document(message.chat.id, open(filename, 'rb'))
    os.remove(filename)

        #list ispolneniya na EOSDO+Pochta
@bot.message_handler(commands=['gliep'])
def get_worker(message):
    doc = Document("LI_Pochta+EOSDO.docx")

    BIGstyle = build_styles(doc)

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
    fio = data[1]
    number = data[2]
    position = data[3]
    department = data[4]
    address = data[5]
    director = data[6]
    date = datetime.today().strftime("%d.%m.%Y")
    name = fio.split(" ") 
    name = name[0] + " " + name[1][0] + "." + (name[2][0] + "." if len(name) > 2 else "")

    apply_style(doc.tables[0].rows[1].cells[1].paragraphs[1], fio, BIGstyle)
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
    
    filename = f"LI_{name}_EOSDO_Pochta.doc"
    doc.save(filename)
    bot.send_document(message.chat.id, open(filename, 'rb'))
    os.remove(filename)

    #list ispolneniya na BOXER
@bot.message_handler(commands=['glib'])
def get_worker(message):
    doc = Document("LI_Boxer.docx")

    BIGstyle = build_styles(doc)

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
    fio = data[1]
    number = data[2]
    position = data[3]
    department = data[4]
    address = data[5]
    director = data[6]
    date = datetime.today().strftime("%d.%m.%Y")
    name = fio.split(" ") 
    name = name[0] + " " + name[1][0] + "." + (name[2][0] + "." if len(name) > 2 else "")

    apply_style(doc.tables[0].rows[1].cells[1].paragraphs[0], fio, BIGstyle)
    apply_style(doc.tables[0].rows[5].cells[1].paragraphs[0], str(number), BIGstyle)
    apply_style(doc.tables[0].rows[6].cells[1].paragraphs[1], position, BIGstyle)
    apply_style(doc.tables[0].rows[8].cells[1].paragraphs[0], department, BIGstyle)
    apply_style(doc.tables[0].rows[9].cells[1].paragraphs[1], address, BIGstyle)
    apply_style(doc.tables[0].rows[9].cells[1].paragraphs[0], '', BIGstyle)
    apply_style(doc.tables[0].rows[14].cells[1].paragraphs[0], name, BIGstyle)
    apply_style(doc.tables[0].rows[14].cells[4].paragraphs[0], date, BIGstyle)
    
    apply_style(doc.tables[2].rows[8].cells[2].paragraphs[0], director, BIGstyle)
    apply_style(doc.tables[2].rows[8].cells[4].paragraphs [0], date, BIGstyle)
    apply_style(doc.tables[2].rows[9].cells[4].paragraphs [0], date, BIGstyle)
    apply_style(doc.tables[2].rows[17].cells[1].paragraphs[0], name, BIGstyle)
    apply_style(doc.tables[2].rows[17].cells[3].paragraphs[0], date, BIGstyle)
    apply_style(doc.tables[2].rows[3].cells[4].paragraphs [0], date, BIGstyle)
    
    filename = f"LI_{name}_BOXER.doc"
    doc.save(filename)
    bot.send_document(message.chat.id, open(filename, 'rb'))
    os.remove(filename)

    # list ispolneniya na 1C eosdo boxer
@bot.message_handler(commands=['gliceb'])
def get_worker(message):
    doc = Document("LI_EOSDO+1C+BOXER.docx")

    BIGstyle = build_styles(doc)

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
    fio = data[1]
    number = data[2]
    position = data[3]
    department = data[4]
    address = data[5]
    director = data[6]
    date = datetime.today().strftime("%d.%m.%Y")
    name = fio.split(" ") 
    name = name[0] + " " + name[1][0] + "." + (name[2][0] + "." if len(name) > 2 else "")

    apply_style(doc.tables[0].rows[1].cells[1].paragraphs[1], fio, BIGstyle)
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

    filename = f"LI_{name}_Eosdo_1C_BOXER.doc"
    doc.save(filename)
    bot.send_document(message.chat.id, open(filename, 'rb'))
    os.remove(filename)

    #list ispolneniya na 1C boxer
@bot.message_handler(commands=['glicb'])
def get_worker(message):
    doc = Document("LI_1C+BOXER.docx")

    BIGstyle = build_styles(doc)

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
    fio = data[1]
    number = data[2]
    position = data[3]
    department = data[4]
    address = data[5]
    director = data[6]
    date = datetime.today().strftime("%d.%m.%Y")
    name = fio.split(" ") 
    name = name[0] + " " + name[1][0] + "." + (name[2][0] + "." if len(name) > 2 else "")

    apply_style(doc.tables[0].rows[1].cells[1].paragraphs[1], fio, BIGstyle)
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
    
    filename = f"LI_{name}_1C_BOXER.doc"
    doc.save(filename)
    bot.send_document(message.chat.id, open(filename, 'rb'))
    os.remove(filename)

     # list ispolneniya na pochta 1c boxer
@bot.message_handler(commands=['glipcb'])
def get_worker(message):
    doc = Document("LI_Pochta+1C+BOXER.docx")

    BIGstyle = build_styles(doc)

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
    fio = data[1]
    number = data[2]
    position = data[3]
    department = data[4]
    address = data[5]
    director = data[6]
    date = datetime.today().strftime("%d.%m.%Y")
    name = fio.split(" ") 
    name = name[0] + " " + name[1][0] + "." + (name[2][0] + "." if len(name) > 2 else "")

    apply_style(doc.tables[0].rows[1].cells[1].paragraphs[1], fio, BIGstyle)
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

    filename = f"LI_{name}1C_BOXER_Pochta.doc"
    doc.save(filename)
    bot.send_document(message.chat.id, open(filename, 'rb'))
    os.remove(filename)

#list ispolneniya na EOSDO+BOxer
@bot.message_handler(commands=['glieb'])
def get_worker(message):
    doc = Document("LI_EOSDO+BOXER.docx")

    BIGstyle = build_styles(doc)

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
    fio = data[1]
    number = data[2]
    position = data[3]
    department = data[4]
    address = data[5]
    director = data[6]
    date = datetime.today().strftime("%d.%m.%Y")
    name = fio.split(" ") 
    name = name[0] + " " + name[1][0] + "." + (name[2][0] + "." if len(name) > 2 else "")

    apply_style(doc.tables[0].rows[1].cells[1].paragraphs[1], fio, BIGstyle)
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
    
    filename = f"LI_{name}_EOSDO_BOXER.doc"
    doc.save(filename)
    bot.send_document(message.chat.id, open(filename, 'rb'))
    os.remove(filename)

    # list ispolneniya na EOSDO Pochta BOXER
@bot.message_handler(commands=['gliepb'])
def get_worker(message):
    doc = Document("LI_Pochta+EOSDO+BOXER.docx")

    BIGstyle = build_styles(doc)

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
    fio = data[1]
    number = data[2]
    position = data[3]
    department = data[4]
    address = data[5]
    director = data[6]
    date = datetime.today().strftime("%d.%m.%Y")
    name = fio.split(" ") 
    name = name[0] + " " + name[1][0] + "." + (name[2][0] + "." if len(name) > 2 else "")

    apply_style(doc.tables[0].rows[1].cells[1].paragraphs[1], fio, BIGstyle)
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

    filename = f"LI_{name}_Eosdo_Pochta_BOXER.doc"
    doc.save(filename)
    bot.send_document(message.chat.id, open(filename, 'rb'))
    os.remove(filename)

    # list ispolneniya na 1C eosdo boxer
@bot.message_handler(commands=['glipce'])
def get_worker(message):
    doc = Document("LI_Pochta+EOSDO+1C.docx")

    BIGstyle = build_styles(doc)

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
    fio = data[1]
    number = data[2]
    position = data[3]
    department = data[4]
    address = data[5]
    director = data[6]
    date = datetime.today().strftime("%d.%m.%Y")
    name = fio.split(" ") 
    name = name[0] + " " + name[1][0] + "." + (name[2][0] + "." if len(name) > 2 else "")

    apply_style(doc.tables[0].rows[1].cells[1].paragraphs[1], fio, BIGstyle)
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

    filename = f"LI_{name}_Eosdo_Pochta_1C.doc"
    doc.save(filename)
    bot.send_document(message.chat.id, open(filename, 'rb'))
    os.remove(filename)


bot.infinity_polling()

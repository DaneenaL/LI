import telebot 

def document_keyboard() -> telebot.types.InlineKeyboardMarkup:
    keyboard = telebot.types.InlineKeyboardMarkup()

    buttons = [
        telebot.types.InlineKeyboardButton(text="Лист исполнения на 1С", callback_data='glic')
    ]
    
    keyboard.add(*buttons)

    return keyboard

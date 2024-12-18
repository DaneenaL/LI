from telebot.types import InlineKeyboardButton as Button, InlineKeyboardMarkup

def document_keyboard() -> InlineKeyboardMarkup:
    keyboard = InlineKeyboardMarkup()

    buttons = [
        Button(text="Лист исполнения на 1С", callback_data='glic')
    ]
    
    keyboard.add(*buttons)

    return keyboard

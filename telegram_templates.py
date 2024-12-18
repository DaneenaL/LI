from telebot.types import InlineKeyboardButton as Button, InlineKeyboardMarkup

def document_keyboard() -> InlineKeyboardMarkup:
    keyboard = InlineKeyboardMarkup()

    buttons = [
        Button(text="1С", callback_data='glic'),
        Button(text="ЕОСДО", callback_data='glie'),
        Button(text="Почту, ЕОСДО, BOXER", callback_data='glipceb'),
        
    ]
    
    keyboard.add(*buttons)

    return keyboard

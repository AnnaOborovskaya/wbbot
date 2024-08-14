from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton

keyb_1 = InlineKeyboardMarkup(row_width=1)
keyb_1.add(InlineKeyboardButton(text='Создать автоматическую кампанию', callback_data='btn_1'))
keyb_1.add(InlineKeyboardButton(text='Создать кампанию Поиск + Каталог (теперь аукцион)', callback_data='btn_2'))
keyb_1.add(InlineKeyboardButton(text='Удалить кампанию', callback_data='btn_3'))
keyb_1.add(InlineKeyboardButton(text='Изменение ставки у кампании', callback_data='btn_4'))

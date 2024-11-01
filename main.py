import openpyxl
import telebot
from telebot.types import InlineKeyboardMarkup, InlineKeyboardButton

bot = telebot.TeleBot("token")

# Загрузка данных о товарах из XLSX файла
products = []
try:
    workbook = openpyxl.load_workbook('main.xlsx')
    sheet = workbook.active
    for row in sheet.iter_rows(min_row=2, values_only=True):
        try:
            product = {
                "type": row[0],
                "name": row[1],
                "description": row[2] or "Описание недоступно",
                "price": row[3] if isinstance(row[3], (int, float)) else "Цена не указана",
                "vel": row[4],
                "data": row[5]
            }
            if product["vel"] == 1:
                products.append(product)
        except Exception as e:
            print(f"Ошибка обработки строки: {e}")
except FileNotFoundError:
    print("Файл не найден. Проверьте путь к файлу.")
except Exception as e:
    print(f"Ошибка при загрузке файла: {e}")

ITEMS_PER_PAGE = 10  # Количество товаров на странице

def main_menu():
    """Создает главное меню."""
    markup = InlineKeyboardMarkup(row_width=1)
    markup.add(InlineKeyboardButton("📋 Посмотреть товары", callback_data="view_products"))
    markup.add(InlineKeyboardButton("🔍 Поиск товаров", callback_data="search_products"))
    markup.add(InlineKeyboardButton("📞 Связаться с продавцом", url="https://t.me/jahamars"))
    return "Выберите действие:", markup





@bot.message_handler(commands=['start'])
def welcome(message):
    bot.send_message(message.chat.id, "Приветствую! Добро пожаловать.")
    text, markup = main_menu()
    bot.send_message(message.chat.id, text, reply_markup=markup)

def get_catalog_page(page=0):
    start_index = page * ITEMS_PER_PAGE
    end_index = start_index + ITEMS_PER_PAGE
    markup = InlineKeyboardMarkup(row_width=2)

    for item in products[start_index:end_index]:
        markup.add(InlineKeyboardButton(item["name"], callback_data=f"item_{item['name']}"))

    # Кнопки навигации
    navigation_buttons = []
    if page > 0:
        navigation_buttons.append(InlineKeyboardButton("⬅️ Назад", callback_data=f"page_{page - 1}"))
    if end_index < len(products):
        navigation_buttons.append(InlineKeyboardButton("➡️ Вперед", callback_data=f"page_{page + 1}"))

    
    for i in range((len(products) // ITEMS_PER_PAGE) + 1):
        navigation_buttons.append(InlineKeyboardButton(f"{i + 1}", callback_data=f"page_{i}"))

    navigation_buttons.append(InlineKeyboardButton("🏠 Главное меню", callback_data="main_menu"))

    markup.add(*navigation_buttons)
    
    return f"Страница {page + 1}", markup

@bot.callback_query_handler(func=lambda call: call.data == "view_products")
def catalog(call):
    
    text, markup = get_catalog_page(0)
    bot.send_message(call.message.chat.id, text, reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data.startswith('page_'))
def navigate(call):
    
    page = int(call.data.split('_')[1])
    text, markup = get_catalog_page(page)
    bot.send_message(call.message.chat.id, text, reply_markup=markup)



#------
@bot.callback_query_handler(func=lambda call: call.data == "search_products")
def search_products(call):
    
    bot.send_message(call.message.chat.id, "Введите название товара для поиска:")
    bot.register_next_step_handler(call.message, perform_search)

def perform_search(message):
    
    query = message.text.lower()
    found_items = [item for item in products if query in item["name"].lower()]
    
    if not found_items:
        bot.send_message(message.chat.id, "Товаров по вашему запросу не найдено.")
        return

    markup = InlineKeyboardMarkup(row_width=2)
    for item in found_items:
        markup.add(InlineKeyboardButton(item["name"], callback_data=f"item_{item['name']}"))

    markup.add(InlineKeyboardButton("🏠 Главное меню", callback_data="main_menu"))
    bot.send_message(message.chat.id, "Результаты поиска:", reply_markup=markup)

#------



@bot.callback_query_handler(func=lambda call: call.data.startswith('item_'))
def show_item(call):
    item_name = call.data.split('_')[1]
    item = next((product for product in products if product["name"] == item_name), None)
    if item:
        description = item.get("description", "Описание недоступно")
        price = item.get("price", "Цена не указана")
        response_text = f"**{item_name}**\n\n{description}\n\nЦена: {price} руб."
        markup = InlineKeyboardMarkup(row_width=1)
        markup.add(
            InlineKeyboardButton("⬅️ Назад к товарам", callback_data="return_to_catalog"),
            InlineKeyboardButton("📦 Заказать", url="https://t.me/jahamars"),
            InlineKeyboardButton("🏠 Главное меню", callback_data="main_menu")
        )

        bot.send_message(call.message.chat.id, response_text, reply_markup=markup)





@bot.callback_query_handler(func=lambda call: call.data == "return_to_catalog")
def return_to_catalog(call):
    bot.delete_message(call.message.chat.id, call.message.message_id)





@bot.callback_query_handler(func=lambda call: call.data == "main_menu")
def return_to_main_menu(call):
    text, markup = main_menu()
    bot.send_message(call.message.chat.id, text, reply_markup=markup)

print("Бот запущен...")
bot.polling()

import asyncio
import logging
import os
import json
from datetime import datetime, time
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage

# ================== НАСТРОЙКИ ==================
TOKEN = os.getenv("BOT_TOKEN", "8873059239:AAF8FMbc7AlxI15PHXUHv_dX9inxKBWAvfY")
ADMIN_IDS = [8725919396]  # Добавь второй ID сюда

WORK_START = time(9, 0)
WORK_END = time(22, 0)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

bot = Bot(token=TOKEN)
storage = MemoryStorage()
dp = Dispatcher(storage=storage)

# ================== ПРОМОКОДЫ ==================
PROMO_CODES = {
    "VR10": 10,
    "VR20": 20,
    "ALMATYVR": 15,
}

# ================== ЦЕНЫ ==================
PRICES = {
    "Meta Quest 3": 1500,
    "Meta Quest 3S": 1200,
    "Pico 4 Ultra": 1800,
    "PlayStation VR2 + PS5": 2500,
}
OPERATOR_PRICE = 6500   # тг/час
TRANSPORT_PRICE = 25000  # тг/выезд
TV_PRICE = 55000         # тг/час

# ================== CRM ==================
CRM_FILE = "crm.json"
CLIENTS_FILE = "clients.json"

def load_crm():
    if os.path.exists(CRM_FILE):
        with open(CRM_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_crm(data):
    with open(CRM_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def load_clients():
    if os.path.exists(CLIENTS_FILE):
        with open(CLIENTS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_clients(data):
    with open(CLIENTS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def add_order(order: dict):
    orders = load_crm()
    order["id"] = len(orders) + 1
    order["status"] = "🆕 Новая"
    order["created_at"] = datetime.now().strftime("%d.%m.%Y %H:%M")
    orders.append(order)
    save_crm(orders)

    # Обновляем базу клиентов
    clients = load_clients()
    uid = str(order["user_id"])
    if uid not in clients:
        clients[uid] = {
            "name": order["name"],
            "phone": order["phone"],
            "username": order.get("username", "нет"),
            "orders": [],
            "total_spent": 0,
        }
    clients[uid]["orders"].append(order["id"])
    clients[uid]["total_spent"] = clients[uid].get("total_spent", 0) + order.get("total", 0)
    save_clients(clients)

    return order["id"]

def update_status(order_id: int, status: str):
    orders = load_crm()
    for o in orders:
        if o["id"] == order_id:
            o["status"] = status
            save_crm(orders)
            return True
    return False

def get_user_lang(user_id: int) -> str:
    clients = load_clients()
    uid = str(user_id)
    if uid in clients:
        return clients[uid].get("lang", "ru")
    return "ru"

def set_user_lang(user_id: int, lang: str):
    clients = load_clients()
    uid = str(user_id)
    if uid not in clients:
        clients[uid] = {}
    clients[uid]["lang"] = lang
    save_clients(clients)

# ================== ТЕКСТЫ ==================
T = {
    "welcome": {
        "ru": "👋 <b>Добро пожаловать в VR PROGRESS!</b>\n\n🎮 Аренда VR-шлемов в Алматы\nMeta Quest 3 • Quest 3S • Pico 4 Ultra • PS VR2\n\nВыберите действие ниже 👇",
        "kz": "👋 <b>VR PROGRESS-ке қош келдіңіз!</b>\n\n🎮 Алматыдағы VR дулыға жалдау\nMeta Quest 3 • Quest 3S • Pico 4 Ultra • PS VR2\n\nТөменде әрекетті таңдаңыз 👇"
    },
    "off_hours": {
        "ru": "⏰ Сейчас мы не работаем.\n\nРабочие часы: 9:00–22:00\nОставьте заявку — ответим утром! 🌅",
        "kz": "⏰ Қазір біз жұмыс істемейміз.\n\nЖұмыс уақыты: 9:00–22:00\nӨтінім қалдырыңыз — таңертең жауап береміз! 🌅"
    },
    "catalog": {
        "ru": "<b>📦 Доступное оборудование:</b>\n\n🕹 <b>Meta Quest 3</b> — 128GB / 512GB\n   Флагман для автономной VR\n\n🕹 <b>Meta Quest 3S</b>\n   Компактный и доступный\n\n🕹 <b>Pico 4 Ultra</b>\n   Премиум с отслеживанием взгляда\n\n🕹 <b>PlayStation VR2 + PS5</b>\n   Эксклюзивные игры PlayStation",
        "kz": "<b>📦 Қол жетімді жабдықтар:</b>\n\n🕹 <b>Meta Quest 3</b> — 128GB / 512GB\n   Автономды VR флагманы\n\n🕹 <b>Meta Quest 3S</b>\n   Компактты және қолжетімді\n\n🕹 <b>Pico 4 Ultra</b>\n   Көз бақылауы бар премиум\n\n🕹 <b>PlayStation VR2 + PS5</b>\n   PlayStation эксклюзивті ойындары"
    },
    "prices": {
        "ru": "<b>💰 Цены на аренду (за час):</b>\n\n🎮 Meta Quest 3 — 1 500 тг/час\n🎮 Meta Quest 3S — 1 200 тг/час\n🎮 Pico 4 Ultra — 1 800 тг/час\n🎮 PS VR2 + PS5 — 2 500 тг/час\n\n👨‍💼 Оператор — 6 500 тг/час\n🚗 Транспорт — 25 000 тг/выезд\n📺 ТВ со стойками — 55 000 тг/час\n\n💳 Оплата: наличными или переводом",
        "kz": "<b>💰 Жалдау бағалары (сағатына):</b>\n\n🎮 Meta Quest 3 — 1 500 тг/сағ\n🎮 Meta Quest 3S — 1 200 тг/сағ\n🎮 Pico 4 Ultra — 1 800 тг/сағ\n🎮 PS VR2 + PS5 — 2 500 тг/сағ\n\n👨‍💼 Оператор — 6 500 тг/сағ\n🚗 Көлік — 25 000 тг/шығу\n📺 Стойкалы ТВ — 55 000 тг/сағ\n\n💳 Төлем: қолма-қол немесе аударым"
    },
    "how_rent": {
        "ru": "<b>📖 Как арендовать:</b>\n\n1️⃣ Выберите шлем\n2️⃣ Нажмите «Оформить аренду»\n3️⃣ Заполните форму\n4️⃣ Менеджер свяжется с вами\n\n📍 Работаем по всей Алматы\n🔒 Залог: 10 000 тг\n⏰ Ежедневно: 9:00–22:00",
        "kz": "<b>📖 Қалай жалдауға болады:</b>\n\n1️⃣ Дулығаны таңдаңыз\n2️⃣ «Жалдауды рәсімдеу» батырмасын басыңыз\n3️⃣ Нысанды толтырыңыз\n4️⃣ Менеджер сізбен байланысады\n\n📍 Алматы бойынша жұмыс істейміз\n🔒 Кепілдік: 10 000 тг\n⏰ Күн сайын: 9:00–22:00"
    },
}

# ================== СОСТОЯНИЯ ==================
class LangForm(StatesGroup):
    choosing = State()

class RentForm(StatesGroup):
    equipment = State()
    quantity = State()
    hours = State()
    tv = State()
    promo = State()
    agree = State()
    deposit_photo = State()
    name = State()
    phone = State()
    address = State()
    comments = State()

class ConsultForm(StatesGroup):
    question = State()

# ================== КЛАВИАТУРЫ ==================
def lang_menu():
    keyboard = [
        [types.KeyboardButton(text="🇷🇺 Русский")],
        [types.KeyboardButton(text="🇰🇿 Қазақша")]
    ]
    return types.ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True)

def main_menu(lang="ru"):
    if lang == "ru":
        keyboard = [
            [types.KeyboardButton(text="📋 Каталог")],
            [types.KeyboardButton(text="💰 Цены")],
            [types.KeyboardButton(text="📖 Как арендовать")],
            [types.KeyboardButton(text="🧠 Консультация")],
            [types.KeyboardButton(text="🚀 Оформить аренду")],
            [types.KeyboardButton(text="🔄 Повторная аренда")],
            [types.KeyboardButton(text="🌐 Сменить язык")]
        ]
    else:
        keyboard = [
            [types.KeyboardButton(text="📋 Каталог")],
            [types.KeyboardButton(text="💰 Бағалар")],
            [types.KeyboardButton(text="📖 Қалай жалдауға")],
            [types.KeyboardButton(text="🧠 Кеңес")],
            [types.KeyboardButton(text="🚀 Жалдауды рәсімдеу")],
            [types.KeyboardButton(text="🔄 Қайта жалдау")],
            [types.KeyboardButton(text="🌐 Тілді өзгерту")]
        ]
    return types.ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True)

def admin_menu():
    keyboard = [
        [types.KeyboardButton(text="📊 Все заявки")],
        [types.KeyboardButton(text="🆕 Новые заявки")],
        [types.KeyboardButton(text="📈 Статистика")],
        [types.KeyboardButton(text="👥 База клиентов")],
        [types.KeyboardButton(text="📥 Выгрузить Excel")],
        [types.KeyboardButton(text="🎁 Промокоды")],
        [types.KeyboardButton(text="🔙 Главное меню")]
    ]
    return types.ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True)

def cancel_menu():
    keyboard = [[types.KeyboardButton(text="❌ Отмена")]]
    return types.ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True)

def equipment_menu():
    keyboard = [
        [types.KeyboardButton(text="Meta Quest 3")],
        [types.KeyboardButton(text="Meta Quest 3S")],
        [types.KeyboardButton(text="Pico 4 Ultra")],
        [types.KeyboardButton(text="PlayStation VR2 + PS5")],
        [types.KeyboardButton(text="❌ Отмена")]
    ]
    return types.ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True)

def yes_no_menu():
    keyboard = [
        [types.KeyboardButton(text="✅ Да")],
        [types.KeyboardButton(text="❌ Нет")],
    ]
    return types.ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True)

def agree_menu():
    keyboard = [
        [types.KeyboardButton(text="✅ Принимаю условия")],
        [types.KeyboardButton(text="❌ Отмена")]
    ]
    return types.ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True)

def status_menu(order_id: int):
    keyboard = [
        [types.InlineKeyboardButton(text="🆕 Новая", callback_data=f"status_{order_id}_🆕 Новая")],
        [types.InlineKeyboardButton(text="⏳ В обработке", callback_data=f"status_{order_id}_⏳ В обработке")],
        [types.InlineKeyboardButton(text="🚗 В пути", callback_data=f"status_{order_id}_🚗 В пути")],
        [types.InlineKeyboardButton(text="✅ Выдана", callback_data=f"status_{order_id}_✅ Выдана")],
        [types.InlineKeyboardButton(text="🏁 Закрыта", callback_data=f"status_{order_id}_🏁 Закрыта")],
        [types.InlineKeyboardButton(text="❌ Отменена", callback_data=f"status_{order_id}_❌ Отменена")],
    ]
    return types.InlineKeyboardMarkup(inline_keyboard=keyboard)

def skip_menu():
    keyboard = [
        [types.KeyboardButton(text="⏭ Пропустить")],
        [types.KeyboardButton(text="❌ Отмена")]
    ]
    return types.ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True)

# ================== ПРОВЕРКА РАБОЧИХ ЧАСОВ ==================
def is_working_hours() -> bool:
    now = datetime.now().time()
    return WORK_START <= now <= WORK_END

# ================== ОТМЕНА ==================
@dp.message(F.text == "❌ Отмена")
async def cancel(message: types.Message, state: FSMContext):
    await state.clear()
    lang = get_user_lang(message.from_user.id)
    await message.answer("❌ Отменено." if lang == "ru" else "❌ Болдырылмады.", reply_markup=main_menu(lang))

# ================== СМЕНА ЯЗЫКА ==================
@dp.message(F.text.in_(["🌐 Сменить язык", "🌐 Тілді өзгерту"]))
async def change_lang(message: types.Message, state: FSMContext):
    await state.set_state(LangForm.choosing)
    await message.answer("🌐 Выберите язык / Тілді таңдаңыз:", reply_markup=lang_menu())

@dp.message(LangForm.choosing)
async def set_lang(message: types.Message, state: FSMContext):
    if "Русский" in message.text:
        set_user_lang(message.from_user.id, "ru")
        await state.clear()
        await message.answer("✅ Язык изменён на русский", reply_markup=main_menu("ru"))
    elif "Қазақша" in message.text:
        set_user_lang(message.from_user.id, "kz")
        await state.clear()
        await message.answer("✅ Тіл қазақшаға өзгертілді", reply_markup=main_menu("kz"))

# ================== START ==================
@dp.message(Command("start"))
async def start(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer("🌐 Выберите язык / Тілді таңдаңыз:", reply_markup=lang_menu())
    await state.set_state(LangForm.choosing)

# ================== КАТАЛОГ ==================
@dp.message(F.text == "📋 Каталог")
async def catalog(message: types.Message):
    lang = get_user_lang(message.from_user.id)
    await message.answer(T["catalog"][lang], parse_mode="HTML")

# ================== ЦЕНЫ ==================
@dp.message(F.text.in_(["💰 Цены", "💰 Бағалар"]))
async def prices(message: types.Message):
    lang = get_user_lang(message.from_user.id)
    await message.answer(T["prices"][lang], parse_mode="HTML")

# ================== КАК АРЕНДОВАТЬ ==================
@dp.message(F.text.in_(["📖 Как арендовать", "📖 Қалай жалдауға"]))
async def how_rent(message: types.Message):
    lang = get_user_lang(message.from_user.id)
    await message.answer(T["how_rent"][lang], parse_mode="HTML")

# ================== КОНСУЛЬТАЦИЯ ==================
@dp.message(F.text.in_(["🧠 Консультация", "🧠 Кеңес"]))
async def consult_start(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    if not is_working_hours():
        await message.answer(T["off_hours"][lang], parse_mode="HTML")
        return
    msg = "🧠 <b>Напишите ваш вопрос:</b>" if lang == "ru" else "🧠 <b>Сұрағыңызды жазыңыз:</b>"
    await message.answer(msg, parse_mode="HTML", reply_markup=cancel_menu())
    await state.set_state(ConsultForm.question)

@dp.message(ConsultForm.question)
async def consult_question(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    user = message.from_user
    admin_text = (
        f"💬 <b>ВОПРОС ОТ КЛИЕНТА</b>\n\n"
        f"👤 {user.full_name} (@{user.username or 'нет'})\n"
        f"🆔 ID: <code>{user.id}</code>\n\n"
        f"❓ {message.text}"
    )
    try:
        for admin_id in ADMIN_IDS:
            await bot.send_message(admin_id, admin_text, parse_mode="HTML")
        reply = "✅ Вопрос получен! Ответим в ближайшее время 🙌" if lang == "ru" else "✅ Сұрақ қабылданды! Жақын арада жауап береміз 🙌"
        await message.answer(reply, reply_markup=main_menu(lang))
    except Exception as e:
        logger.error(f"Ошибка: {e}")
    await state.clear()

# ================== ПОВТОРНАЯ АРЕНДА ==================
@dp.message(F.text.in_(["🔄 Повторная аренда", "🔄 Қайта жалдау"]))
async def repeat_rent(message: types.Message):
    lang = get_user_lang(message.from_user.id)
    clients = load_clients()
    uid = str(message.from_user.id)
    if uid not in clients or not clients[uid].get("orders"):
        msg = "📭 У вас пока нет истории аренд." if lang == "ru" else "📭 Сізде жалдау тарихы жоқ."
        await message.answer(msg, reply_markup=main_menu(lang))
        return
    orders = load_crm()
    client = clients[uid]
    last_order_id = client["orders"][-1]
    last = next((o for o in orders if o["id"] == last_order_id), None)
    if not last:
        await message.answer("📭 Заявка не найдена.", reply_markup=main_menu(lang))
        return
    text = (
        f"🔄 <b>Ваша последняя аренда:</b>\n\n"
        f"🎮 {last['equipment']}\n"
        f"⏰ {last.get('hours', '—')} час(ов)\n"
        f"📺 ТВ: {'Да' if last.get('tv') else 'Нет'}\n\n"
        f"Хотите повторить?"
    )
    keyboard = [
        [types.InlineKeyboardButton(text="✅ Повторить", callback_data=f"repeat_{last_order_id}")],
        [types.InlineKeyboardButton(text="❌ Нет", callback_data="repeat_cancel")]
    ]
    await message.answer(text, parse_mode="HTML", reply_markup=types.InlineKeyboardMarkup(inline_keyboard=keyboard))

@dp.callback_query(F.data.startswith("repeat_"))
async def repeat_callback(callback: types.CallbackQuery):
    lang = get_user_lang(callback.from_user.id)
    if callback.data == "repeat_cancel":
        await callback.answer()
        await callback.message.edit_text("❌ Отменено.")
        return
    order_id = int(callback.data.split("_")[1])
    orders = load_crm()
    last = next((o for o in orders if o["id"] == order_id), None)
    if not last:
        await callback.answer("Заявка не найдена.")
        return
    new_order = {
        "equipment": last["equipment"],
        "hours": last.get("hours", 1),
        "tv": last.get("tv", False),
        "name": last["name"],
        "phone": last["phone"],
        "address": last.get("address", "—"),
        "comments": "Повторная аренда",
        "user_id": callback.from_user.id,
        "username": callback.from_user.username or "нет",
        "promo": None,
        "discount": 0,
        "total": last.get("total", 0),
    }
    new_id = add_order(new_order)
    admin_text = f"🔄 <b>ПОВТОРНАЯ ЗАЯВКА #{new_id}</b>\n\n" + format_order({**new_order, "id": new_id, "status": "🆕 Новая", "created_at": datetime.now().strftime("%d.%m.%Y %H:%M")})
    for admin_id in ADMIN_IDS:
        await bot.send_message(admin_id, admin_text, parse_mode="HTML", reply_markup=status_menu(new_id))
    await callback.answer()
    await callback.message.edit_text(f"✅ Повторная заявка #{new_id} отправлена! Менеджер свяжется с вами.")

# ================== ОФОРМЛЕНИЕ ==================
@dp.message(F.text.in_(["🚀 Оформить аренду", "🚀 Жалдауды рәсімдеу"]))
async def rent_start(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    if not is_working_hours():
        await message.answer(T["off_hours"][lang], parse_mode="HTML")
        return
    msg = "🎮 <b>Шаг 1/9</b> — Выберите оборудование:" if lang == "ru" else "🎮 <b>Қадам 1/9</b> — Жабдықты таңдаңыз:"
    await message.answer(msg, parse_mode="HTML", reply_markup=equipment_menu())
    await state.set_state(RentForm.equipment)

@dp.message(RentForm.equipment)
async def rent_equipment(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    await state.update_data(equipment=message.text)
    msg = "🔢 <b>Шаг 2/9</b> — Сколько шлемов нужно?\n\n<i>Введите число (например: 2)</i>" if lang == "ru" else "🔢 <b>Қадам 2/9</b> — Қанша дулыға керек?\n\n<i>Санды енгізіңіз (мысалы: 2)</i>"
    await message.answer(msg, parse_mode="HTML", reply_markup=cancel_menu())
    await state.set_state(RentForm.quantity)

@dp.message(RentForm.quantity)
async def rent_quantity(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    if not message.text.isdigit() or int(message.text) <= 0:
        await message.answer("⚠️ Введите число больше 0" if lang == "ru" else "⚠️ 0-ден үлкен санды енгізіңіз")
        return
    await state.update_data(quantity=int(message.text))
    msg = "⏰ <b>Шаг 3/9</b> — На сколько часов?\n\n<i>Введите число (например: 3)</i>" if lang == "ru" else "⏰ <b>Қадам 3/9</b> — Қанша сағатқа?\n\n<i>Санды енгізіңіз (мысалы: 3)</i>"
    await message.answer(msg, parse_mode="HTML", reply_markup=cancel_menu())
    await state.set_state(RentForm.hours)

@dp.message(RentForm.hours)
async def rent_hours(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    if not message.text.isdigit() or int(message.text) <= 0:
        await message.answer("⚠️ Введите количество часов числом" if lang == "ru" else "⚠️ Сағат санын енгізіңіз")
        return
    await state.update_data(hours=int(message.text))
    msg = "📺 <b>Шаг 4/9</b> — Нужен ТВ со стойками?\n\n💰 55 000 тг/час" if lang == "ru" else "📺 <b>Қадам 4/9</b> — Стойкалы ТВ керек пе?\n\n💰 55 000 тг/сағ"
    await message.answer(msg, parse_mode="HTML", reply_markup=yes_no_menu())
    await state.set_state(RentForm.tv)

@dp.message(RentForm.tv)
async def rent_tv(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    await state.update_data(tv=message.text == "✅ Да")
    msg = "🎁 <b>Шаг 5/9</b> — Есть промокод?\n\nЕсли нет — нажмите «Пропустить»" if lang == "ru" else "🎁 <b>Қадам 5/9</b> — Промокод бар ма?\n\nЖоқ болса — «Өткізу» батырмасын басыңыз"
    await message.answer(msg, parse_mode="HTML", reply_markup=skip_menu())
    await state.set_state(RentForm.promo)

@dp.message(RentForm.promo)
async def rent_promo(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    if message.text == "⏭ Пропустить":
        await state.update_data(promo=None, discount=0)
    else:
        code = message.text.strip().upper()
        if code in PROMO_CODES:
            discount = PROMO_CODES[code]
            await state.update_data(promo=code, discount=discount)
            msg = f"✅ Промокод применён! Скидка {discount}%" if lang == "ru" else f"✅ Промокод қолданылды! Жеңілдік {discount}%"
            await message.answer(msg)
        else:
            await state.update_data(promo=None, discount=0)
            msg = "❌ Промокод не найден. Продолжаем без скидки." if lang == "ru" else "❌ Промокод табылмады. Жеңілдіксіз жалғастырамыз."
            await message.answer(msg)

    msg = "✍️ <b>Шаг 6/9</b> — Ознакомьтесь с условиями:\n\n• Залог: 10 000 тг\n• Бережное обращение с оборудованием\n• Оплата до выдачи\n• Возврат в оговоренное время" if lang == "ru" else "✍️ <b>Қадам 6/9</b> — Шарттармен танысыңыз:\n\n• Кепілдік: 10 000 тг\n• Жабдықты ұқыпты пайдалану\n• Беруден бұрын төлем\n• Келісілген уақытта қайтару"
    await message.answer(msg, parse_mode="HTML", reply_markup=agree_menu())
    await state.set_state(RentForm.agree)

@dp.message(RentForm.agree)
async def rent_agree(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    if message.text != "✅ Принимаю условия":
        await message.answer("❌ Необходимо принять условия для продолжения." if lang == "ru" else "❌ Жалғастыру үшін шарттарды қабылдау керек.")
        return
    msg = "📸 <b>Шаг 7/9</b> — Отправьте фото документа (удостоверение/паспорт)\n\nЭто необходимо для залога.\nЕсли хотите пропустить — нажмите «Пропустить»" if lang == "ru" else "📸 <b>Қадам 7/9</b> — Құжат фотосын жіберіңіз (куәлік/паспорт)\n\nБұл кепілдік үшін қажет.\nӨткізгіңіз келсе — «Өткізу» батырмасын басыңыз"
    await message.answer(msg, parse_mode="HTML", reply_markup=skip_menu())
    await state.set_state(RentForm.deposit_photo)

@dp.message(RentForm.deposit_photo)
async def rent_deposit(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    if message.photo:
        await state.update_data(has_deposit_photo=True)
    else:
        await state.update_data(has_deposit_photo=False)
    msg = "👤 <b>Шаг 8/9</b> — Как вас зовут?" if lang == "ru" else "👤 <b>Қадам 8/9</b> — Есіміңіз кім?"
    await message.answer(msg, parse_mode="HTML", reply_markup=cancel_menu())
    await state.set_state(RentForm.name)

@dp.message(RentForm.name)
async def rent_name(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    await state.update_data(name=message.text)
    msg = "📱 <b>Шаг 9/9</b> — Введите номер телефона:" if lang == "ru" else "📱 <b>Қадам 9/9</b> — Телефон нөміріңізді енгізіңіз:"
    await message.answer(msg, parse_mode="HTML", reply_markup=cancel_menu())
    await state.set_state(RentForm.phone)

@dp.message(RentForm.phone)
async def rent_phone(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    await state.update_data(phone=message.text)
    msg = "📍 Укажите адрес доставки (Алматы):" if lang == "ru" else "📍 Жеткізу мекенжайын көрсетіңіз (Алматы):"
    await message.answer(msg, reply_markup=cancel_menu())
    await state.set_state(RentForm.address)

@dp.message(RentForm.address)
async def rent_address(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    await state.update_data(address=message.text)
    msg = "💬 Дополнительные пожелания? (или напишите «нет»)" if lang == "ru" else "💬 Қосымша тілектер? (немесе «жоқ» деп жазыңыз)"
    await message.answer(msg, reply_markup=cancel_menu())
    await state.set_state(RentForm.comments)

@dp.message(RentForm.comments)
async def rent_finish(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    await state.update_data(comments=message.text)
    data = await state.get_data()

    equipment = data['equipment']
    quantity = data.get('quantity', 1)
    hours = data['hours']
    tv = data.get('tv', False)
    discount = data.get('discount', 0)
    promo = data.get('promo')

    price_per_hour = PRICES.get(equipment, 0)
    equipment_total = price_per_hour * quantity * hours
    operator_total = OPERATOR_PRICE * quantity * hours
    tv_total = TV_PRICE * hours if tv else 0
    subtotal = equipment_total + operator_total + TRANSPORT_PRICE + tv_total
    discount_amount = int(subtotal * discount / 100)
    total = subtotal - discount_amount

    user = message.from_user

    order = {
        "equipment": equipment,
        "quantity": quantity,
        "hours": hours,
        "tv": tv,
        "promo": promo,
        "discount": discount,
        "equipment_cost": equipment_total,
        "operator_cost": operator_total,
        "transport_cost": TRANSPORT_PRICE,
        "tv_cost": tv_total,
        "total": total,
        "name": data['name'],
        "phone": data['phone'],
        "address": data['address'],
        "comments": data['comments'],
        "user_id": user.id,
        "username": user.username or "нет",
        "has_deposit_photo": data.get('has_deposit_photo', False),
        "lang": lang,
    }

    order_id = add_order(order)

    admin_text = (
        f"🔥 <b>НОВАЯ ЗАЯВКА #{order_id}</b>\n\n"
        f"🎮 Оборудование: {equipment} x{quantity}\n"
        f"⏰ Часов: {hours}\n"
        f"📺 ТВ со стойками: {'Да' if tv else 'Нет'}\n\n"
        f"💰 <b>Расчёт:</b>\n"
        f"   Шлемы: {equipment_total:,} тг\n"
        f"   Операторы: {operator_total:,} тг\n"
        f"   Транспорт: {TRANSPORT_PRICE:,} тг\n"
        f"   ТВ: {tv_total:,} тг\n"
        f"   Промокод: {promo or 'нет'} (-{discount}%)\n"
        f"   ─────────────\n"
        f"   Итого: <b>{total:,} тг</b>\n\n"
        f"👤 {data['name']} | 📱 {data['phone']}\n"
        f"📍 {data['address']}\n"
        f"💬 {data['comments']}\n"
        f"📸 Фото залога: {'Да ✅' if data.get('has_deposit_photo') else 'Нет ❌'}\n\n"
        f"Telegram: @{user.username or 'нет'} | ID: <code>{user.id}</code>"
    )

    try:
        for admin_id in ADMIN_IDS:
            await bot.send_message(admin_id, admin_text, parse_mode="HTML", reply_markup=status_menu(order_id))
    except Exception as e:
        logger.error(f"Ошибка отправки заявки: {e}")

    if lang == "ru":
        confirm_text = (
            f"✅ <b>Заявка #{order_id} принята!</b>\n\n"
            f"🎮 {equipment} x{quantity}\n"
            f"⏰ {hours} час(ов)\n"
            f"📺 ТВ: {'Да' if tv else 'Нет'}\n\n"
            f"💰 <b>Итого: {total:,} тг</b>\n\n"
            f"Менеджер свяжется с вами в течение часа 🚀"
        )
    else:
        confirm_text = (
            f"✅ <b>Өтінім #{order_id} қабылданды!</b>\n\n"
            f"🎮 {equipment} x{quantity}\n"
            f"⏰ {hours} сағат\n"
            f"📺 ТВ: {'Иә' if tv else 'Жоқ'}\n\n"
            f"💰 <b>Барлығы: {total:,} тг</b>\n\n"
            f"Менеджер бір сағат ішінде хабарласады 🚀"
        )

    await message.answer(confirm_text, parse_mode="HTML", reply_markup=main_menu(lang))
    await state.clear()

# ================== СМЕНА СТАТУСА ==================
@dp.callback_query(F.data.startswith("status_"))
async def change_status_cb(callback: types.CallbackQuery):
    if callback.from_user.id not in ADMIN_IDS:
        return
    parts = callback.data.split("_", 2)
    order_id = int(parts[1])
    new_status = parts[2]
    update_status(order_id, new_status)
    await callback.answer(f"Статус: {new_status}")
    try:
        await callback.message.edit_text(
            callback.message.text + f"\n\n✏️ Статус: <b>{new_status}</b>",
            parse_mode="HTML"
        )
    except:
        pass

# ================== АДМИН ==================
@dp.message(Command("admin"))
async def admin_panel(message: types.Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    await message.answer("👑 <b>Панель администратора</b>", parse_mode="HTML", reply_markup=admin_menu())

@dp.message(F.text == "🔙 Главное меню")
async def back_main(message: types.Message):
    lang = get_user_lang(message.from_user.id)
    await message.answer("Главное меню", reply_markup=main_menu(lang))

@dp.message(F.text == "📊 Все заявки")
async def all_orders(message: types.Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    orders = load_crm()
    if not orders:
        await message.answer("📭 Заявок нет.")
        return
    for o in orders[-10:][::-1]:
        await message.answer(format_order(o), parse_mode="HTML", reply_markup=status_menu(o["id"]))

@dp.message(F.text == "🆕 Новые заявки")
async def new_orders(message: types.Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    orders = [o for o in load_crm() if o["status"] == "🆕 Новая"]
    if not orders:
        await message.answer("✅ Новых заявок нет.")
        return
    for o in orders:
        await message.answer(format_order(o), parse_mode="HTML", reply_markup=status_menu(o["id"]))

@dp.message(F.text == "📈 Статистика")
async def statistics(message: types.Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    orders = load_crm()
    total = len(orders)
    new = len([o for o in orders if o["status"] == "🆕 Новая"])
    in_progress = len([o for o in orders if "обработке" in o["status"]])
    closed = len([o for o in orders if o["status"] == "🏁 Закрыта"])
    revenue = sum(o.get("total", 0) for o in orders if o["status"] == "🏁 Закрыта")

    today = datetime.now().strftime("%d.%m.%Y")
    today_orders = [o for o in orders if o.get("created_at", "").startswith(today)]

    text = (
        f"📈 <b>Статистика VR PROGRESS</b>\n\n"
        f"📋 Всего заявок: <b>{total}</b>\n"
        f"📅 Сегодня: <b>{len(today_orders)}</b>\n"
        f"🆕 Новых: <b>{new}</b>\n"
        f"⏳ В обработке: <b>{in_progress}</b>\n"
        f"🏁 Закрытых: <b>{closed}</b>\n\n"
        f"💰 Выручка (закрытые): <b>{revenue:,} тг</b>"
    )
    await message.answer(text, parse_mode="HTML")

@dp.message(F.text == "👥 База клиентов")
async def client_base(message: types.Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    clients = load_clients()
    if not clients:
        await message.answer("📭 База клиентов пуста.")
        return
    text = f"👥 <b>База клиентов ({len(clients)} чел.):</b>\n\n"
    for uid, c in list(clients.items())[:20]:
        text += (
            f"👤 {c.get('name', '—')} | 📱 {c.get('phone', '—')}\n"
            f"   @{c.get('username', 'нет')} | Заявок: {len(c.get('orders', []))} | Потрачено: {c.get('total_spent', 0):,} тг\n\n"
        )
    await message.answer(text, parse_mode="HTML")

@dp.message(F.text == "🎁 Промокоды")
async def show_promos(message: types.Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    text = "<b>🎁 Активные промокоды:</b>\n\n"
    for code, discount in PROMO_CODES.items():
        text += f"• <code>{code}</code> — скидка {discount}%\n"
    await message.answer(text, parse_mode="HTML")

@dp.message(F.text == "📥 Выгрузить Excel")
async def export_excel(message: types.Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    orders = load_crm()
    if not orders:
        await message.answer("📭 Нет данных для выгрузки.")
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Заявки"

    headers = ["#", "Дата", "Статус", "Оборудование", "Кол-во", "Часов", "ТВ", "Оператор", "Транспорт", "Промокод", "Скидка%", "Итого", "Имя", "Телефон", "Адрес", "Комментарий", "Telegram"]
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for row, o in enumerate(orders, 2):
        ws.cell(row=row, column=1, value=o.get("id"))
        ws.cell(row=row, column=2, value=o.get("created_at"))
        ws.cell(row=row, column=3, value=o.get("status"))
        ws.cell(row=row, column=4, value=o.get("equipment"))
        ws.cell(row=row, column=5, value=o.get("quantity", 1))
        ws.cell(row=row, column=6, value=o.get("hours"))
        ws.cell(row=row, column=7, value="Да" if o.get("tv") else "Нет")
        ws.cell(row=row, column=8, value=o.get("operator_cost", 0))
        ws.cell(row=row, column=9, value=o.get("transport_cost", 0))
        ws.cell(row=row, column=10, value=o.get("promo") or "—")
        ws.cell(row=row, column=11, value=o.get("discount", 0))
        ws.cell(row=row, column=12, value=o.get("total", 0))
        ws.cell(row=row, column=13, value=o.get("name"))
        ws.cell(row=row, column=14, value=o.get("phone"))
        ws.cell(row=row, column=15, value=o.get("address"))
        ws.cell(row=row, column=16, value=o.get("comments"))
        ws.cell(row=row, column=17, value=f"@{o.get('username', 'нет')}")

    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 18

    filename = f"vrprogress_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx"
    wb.save(filename)

    with open(filename, "rb") as f:
        await message.answer_document(types.BufferedInputFile(f.read(), filename=filename), caption="📥 Выгрузка заявок VR PROGRESS")
    os.remove(filename)

# ================== УТРЕННИЙ ОТЧЁТ ==================
async def morning_report():
    while True:
        now = datetime.now()
        if now.hour == 9 and now.minute == 0:
            orders = load_crm()
            today = now.strftime("%d.%m.%Y")
            yesterday = orders  # все заявки за вчера
            new_count = len([o for o in orders if o["status"] == "🆕 Новая"])
            total_count = len(orders)
            revenue = sum(o.get("total", 0) for o in orders if o["status"] == "🏁 Закрыта")

            text = (
                f"🌅 <b>Доброе утро! Отчёт на {today}</b>\n\n"
                f"📋 Всего заявок: <b>{total_count}</b>\n"
                f"🆕 Необработанных: <b>{new_count}</b>\n"
                f"💰 Общая выручка: <b>{revenue:,} тг</b>\n\n"
                f"Удачного дня! 💪"
            )
            for admin_id in ADMIN_IDS:
                try:
                    await bot.send_message(admin_id, text, parse_mode="HTML")
                except Exception as e:
                    logger.error(f"Ошибка отчёта: {e}")
            await asyncio.sleep(60)
        await asyncio.sleep(30)

# ================== ВСПОМОГАТЕЛЬНЫЕ ==================
def format_order(o: dict) -> str:
    return (
        f"📋 <b>Заявка #{o['id']}</b> | {o.get('status', '—')}\n"
        f"🕐 {o.get('created_at', '—')}\n\n"
        f"🎮 {o.get('equipment')} x{o.get('quantity', 1)}\n"
        f"⏰ Часов: {o.get('hours', '—')}\n"
        f"📺 ТВ: {'Да' if o.get('tv') else 'Нет'}\n"
        f"💰 Итого: {o.get('total', 0):,} тг\n"
        f"👤 {o.get('name')} | 📱 {o.get('phone')}\n"
        f"📍 {o.get('address', '—')}\n"
        f"💬 {o.get('comments', '—')}\n"
        f"Telegram: @{o.get('username', 'нет')}"
    )

# ================== ЗАПУСК ==================
async def main():
    logger.info("🤖 Бот запущен")
    asyncio.create_task(morning_report())
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())

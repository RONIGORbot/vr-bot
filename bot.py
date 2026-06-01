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
ADMIN_IDS = [8725919396, 294017914]

WORK_START = time(9, 0)
WORK_END = time(22, 0)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

bot = Bot(token=TOKEN)
storage = MemoryStorage()
dp = Dispatcher(storage=storage)

# ================== ПРОМОКОДЫ ==================
PROMO_CODES = {}

# ================== ЦЕНЫ ==================
PRICE_STANDARD = 8500    # тг/час за 1 шлем (до 4 шт.)
PRICE_BULK = 7500        # тг/час при 5+ шлемах
MIN_HOURS = 3            # минимальный заказ
OPERATOR_PRICE = 6500    # тг/час на оператора
TRANSPORT_PRICE = 25000  # тг/выезд (фикс)
TV_PRICE = 10500         # тг/час за ТВ

# Медиа-пакеты (цена за мероприятие, от указанной суммы)
MEDIA_PACKAGES = {
    "🎬 Кинокамера Sony FX30 Cinema Line форматы (съёмка + монтаж Reels)": 20000,
    "📱 Мобилография (iPhone Pro Max)": 20000,
    "🎥 Полный пакет (кинокамера + мобилография)": 50000,
    "❌ Без медиа-сопровождения": 0,
}

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

    clients = load_clients()
    uid = str(order["user_id"])
    if uid not in clients:
        clients[uid] = {}
    clients[uid].setdefault("orders", [])
    clients[uid].setdefault("total_spent", 0)
    clients[uid]["name"] = order["name"]
    clients[uid]["phone"] = order["phone"]
    clients[uid]["username"] = order.get("username", "нет")
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
        "ru": (
            "👋 <b>Добро пожаловать в VR PROGRESS!</b>\n\n"
            "🏆 <i>Компания с 9-летним опытом в VR-индустрии</i>\n\n"
            "Мы организуем VR-мероприятия под ключ: от оборудования Meta Quest 3 "
            "до профессионального видеосопровождения (кинокамера / мобилография). "
            "Выполним любую задачу.\n\n"
            "🎮 Meta Quest 3\n\n"
            "Выберите действие ниже 👇"
        ),
        "kz": (
            "👋 <b>VR PROGRESS-ке қош келдіңіз!</b>\n\n"
            "🏆 <i>VR индустриясындағы 9 жылдық тәжірибесі бар компания</i>\n\n"
            "Біз Meta Quest 3 жабдығынан бастап кәсіби бейне сүйемелдеуге дейін "
            "VR іс-шараларын кілт тапсыру арқылы ұйымдастырамыз. "
            "Кез келген тапсырманы орындаймыз.\n\n"
            "🎮 Meta Quest 3\n\n"
            "Төменде әрекетті таңдаңыз 👇"
        ),
    },
    "off_hours": {
        "ru": "⏰ Сейчас мы не работаем.\n\nРабочие часы: 9:00–22:00\nОставьте заявку — ответим утром! 🌅",
        "kz": "⏰ Қазір біз жұмыс істемейміз.\n\nЖұмыс уақыты: 9:00–22:00\nӨтінім қалдырыңыз — таңертең жауап береміз! 🌅",
    },
    "catalog": {
        "ru": (
            "<b>📦 Доступное оборудование:</b>\n\n"
            "🕹 <b>Meta Quest 3</b> — 128GB / 512GB\n"
            "   Флагман для автономной VR"
        ),
        "kz": (
            "<b>📦 Қол жетімді жабдықтар:</b>\n\n"
            "🕹 <b>Meta Quest 3</b> — 128GB / 512GB\n"
            "   Автономды VR флагманы"
        ),
    },
    "prices": {
        "ru": (
            "<b>💰 Базовые тарифы (за 1 час):</b>\n\n"
            "🎮 VR-шлем Meta Quest 3 — 8 500 тг/час\n"
            "   🔥 От 5 шлемов — 7 500 тг/час\n"
            "   ⏱ Минимальный заказ — от 3 часов\n\n"
            "👨‍💼 Оператор/инструктор — 6 500 тг/час\n"
            "📺 Телевизор (Full HD/4K) — 10 500 тг/час\n"
            "🚗 Логистика (доставка + монтаж) — 25 000 тг/выезд\n\n"
            "<b>🎬 Медиа-сопровождение (за мероприятие):</b>\n"
            "   🎬 Кинокамера Sony FX30 Cinema Line (съёмка + монтаж Reels) — от 20 000 тг\n"
            "   📱 Мобилография (iPhone Pro Max) — от 20 000 тг\n"
            "   🎥 Полный пакет (оба варианта) — от 35 000 тг\n\n"
            "💳 Оплата: наличными или переводом"
        ),
        "kz": (
            "<b>💰 Негізгі тарифтер (сағатына):</b>\n\n"
            "🎮 VR дулығасы Meta Quest 3 — 8 500 тг/сағ\n"
            "   🔥 5 дулығадан — 7 500 тг/сағ\n"
            "   ⏱ Ең аз тапсырыс — 3 сағаттан\n\n"
            "👨‍💼 Оператор/нұсқаушы — 6 500 тг/сағ\n"
            "📺 Теледидар (Full HD/4K) — 10 500 тг/сағ\n"
            "🚗 Логистика (жеткізу + монтаж) — 25 000 тг/шығу\n\n"
            "<b>🎬 Медиа-сүйемелдеу (іс-шараға):</b>\n"
            "   🎬 Кинокамера Sony FX30 Cinema Line (түсіру + монтаж) — 20 000 тг-дан\n"
            "   📱 Мобилография (iPhone Pro Max) — 20 000 тг-дан\n"
            "   🎥 Толық пакет (екі нұсқа) — 35 000 тг-дан\n\n"
            "💳 Төлем: қолма-қол немесе аударым"
        ),
    },
    "how_rent": {
        "ru": (
            "<b>📖 Как арендовать:</b>\n\n"
            "1️⃣ Выберите шлем\n"
            "2️⃣ Нажмите «Оформить аренду»\n"
            "3️⃣ Заполните форму\n"
            "4️⃣ Менеджер свяжется с вами\n\n"
            "📍 Работаем по всей Алматы\n"
            "⏰ Ежедневно: 9:00–22:00"
        ),
        "kz": (
            "<b>📖 Қалай жалдауға болады:</b>\n\n"
            "1️⃣ Дулығаны таңдаңыз\n"
            "2️⃣ «Жалдауды рәсімдеу» батырмасын басыңыз\n"
            "3️⃣ Нысанды толтырыңыз\n"
            "4️⃣ Менеджер сізбен байланысады\n\n"
            "📍 Алматы бойынша жұмыс істейміз\n"
            "⏰ Күн сайын: 9:00–22:00"
        ),
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
    tv_quantity = State()
    operators = State()
    media = State()
    name = State()
    phone = State()
    address = State()
    comments = State()

class ConsultForm(StatesGroup):
    question = State()

class CameraForm(StatesGroup):
    hours = State()
    name = State()
    phone = State()
    address = State()
    comments = State()

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
            [types.KeyboardButton(text="🎬 Аренда кинокамеры")],
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
            [types.KeyboardButton(text="🎬 Кинокамера жалдау")],
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

def media_menu():
    keyboard = [
        [types.KeyboardButton(text="🎬 Кинокамера Sony FX30 Cinema Line (съёмка + монтаж Reels)")],
        [types.KeyboardButton(text="📱 Мобилография (iPhone Pro Max)")],
        [types.KeyboardButton(text="🎥 Полный пакет (кинокамера + мобилография)")],
        [types.KeyboardButton(text="❌ Без медиа-сопровождения")],
        [types.KeyboardButton(text="❌ Отмена")]
    ]
    return types.ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True)

def confirm_order_menu():
    keyboard = [
        [types.KeyboardButton(text="✅ Оформить заявку")],
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
        "hours": last.get("hours", 3),
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

# ================== ОФОРМЛЕНИЕ АРЕНДЫ ==================
@dp.message(F.text.in_(["🚀 Оформить аренду", "🚀 Жалдауды рәсімдеу"]))
async def rent_start(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    if not is_working_hours():
        await message.answer(T["off_hours"][lang], parse_mode="HTML")
        return
    msg = "🎮 <b>Шаг 1/8</b> — Выберите оборудование:" if lang == "ru" else "🎮 <b>Қадам 1/8</b> — Жабдықты таңдаңыз:"
    await message.answer(msg, parse_mode="HTML", reply_markup=equipment_menu())
    await state.set_state(RentForm.equipment)

@dp.message(RentForm.equipment)
async def rent_equipment(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    valid = ["Meta Quest 3"]
    if message.text not in valid:
        await message.answer("⚠️ Выберите оборудование из списка" if lang == "ru" else "⚠️ Тізімнен жабдықты таңдаңыз")
        return
    await state.update_data(equipment=message.text)
    msg = (
        f"🔢 <b>Шаг 2/8</b> — Сколько шлемов нужно?\n\n"
        f"<i>Введите число (мин. 1)\n🔥 От 5 шт. — оптовая цена 7 500 тг/час</i>"
    ) if lang == "ru" else (
        f"🔢 <b>Қадам 2/8</b> — Қанша дулыға керек?\n\n"
        f"<i>Санды енгізіңіз (мин. 1)\n🔥 5 дана-дан — жоғарыда баға 7 500 тг/сағ</i>"
    )
    await message.answer(msg, parse_mode="HTML", reply_markup=cancel_menu())
    await state.set_state(RentForm.quantity)

@dp.message(RentForm.quantity)
async def rent_quantity(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    if not message.text.isdigit() or int(message.text) <= 0:
        await message.answer("⚠️ Введите число больше 0" if lang == "ru" else "⚠️ 0-ден үлкен санды енгізіңіз")
        return
    qty = int(message.text)
    await state.update_data(quantity=qty)
    bulk_note = ""
    if qty >= 5:
        bulk_note = "\n\n🔥 <b>Применена оптовая цена 7 500 тг/час!</b>" if lang == "ru" else "\n\n🔥 <b>Жоғарыда баға 7 500 тг/сағ қолданылды!</b>"
    msg = (
        f"⏰ <b>Шаг 3/8</b> — На сколько часов?{bulk_note}\n\n"
        f"<i>Минимальный заказ — {MIN_HOURS} часа</i>"
    ) if lang == "ru" else (
        f"⏰ <b>Қадам 3/8</b> — Қанша сағатқа?{bulk_note}\n\n"
        f"<i>Ең аз тапсырыс — {MIN_HOURS} сағат</i>"
    )
    await message.answer(msg, parse_mode="HTML", reply_markup=cancel_menu())
    await state.set_state(RentForm.hours)

@dp.message(RentForm.hours)
async def rent_hours(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    if not message.text.isdigit() or int(message.text) <= 0:
        await message.answer("⚠️ Введите количество часов числом" if lang == "ru" else "⚠️ Сағат санын енгізіңіз")
        return
    hours = int(message.text)
    if hours < MIN_HOURS:
        warn = f"⚠️ Минимальный заказ — {MIN_HOURS} часа. Введите {MIN_HOURS} или больше." if lang == "ru" else f"⚠️ Ең аз тапсырыс — {MIN_HOURS} сағат. {MIN_HOURS} немесе көп санды енгізіңіз."
        await message.answer(warn)
        return
    await state.update_data(hours=hours)
    msg = (
        f"📺 <b>Шаг 4/8</b> — Нужен телевизор (Full HD/4K трансляция)?\n\n"
        f"💰 {TV_PRICE:,} тг/час"
    ) if lang == "ru" else (
        f"📺 <b>Қадам 4/8</b> — Теледидар керек пе (Full HD/4K трансляция)?\n\n"
        f"💰 {TV_PRICE:,} тг/сағ"
    )
    await message.answer(msg, parse_mode="HTML", reply_markup=yes_no_menu())
    await state.set_state(RentForm.tv)

@dp.message(RentForm.tv)
async def rent_tv(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    need_tv = message.text == "✅ Да"
    await state.update_data(tv=need_tv)
    if need_tv:
        msg = "📺 <b>Шаг 5/8</b> — Сколько телевизоров нужно?\n\n<i>Введите число (например: 1)</i>" if lang == "ru" else "📺 <b>Қадам 5/8</b> — Қанша теледидар керек?\n\n<i>Санды енгізіңіз (мысалы: 1)</i>"
        await message.answer(msg, parse_mode="HTML", reply_markup=cancel_menu())
        await state.set_state(RentForm.tv_quantity)
    else:
        await state.update_data(tv_quantity=0)
        msg = (
            f"👨‍💼 <b>Шаг 5/8</b> — Сколько операторов/инструкторов?\n\n"
            f"💰 {OPERATOR_PRICE:,} тг/час\n<i>Введите 0, если не нужны</i>"
        ) if lang == "ru" else (
            f"👨‍💼 <b>Қадам 5/8</b> — Қанша оператор/нұсқаушы керек?\n\n"
            f"💰 {OPERATOR_PRICE:,} тг/сағ\n<i>Қажет болмаса 0 енгізіңіз</i>"
        )
        await message.answer(msg, parse_mode="HTML", reply_markup=cancel_menu())
        await state.set_state(RentForm.operators)

@dp.message(RentForm.tv_quantity)
async def rent_tv_quantity(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    if not message.text.isdigit() or int(message.text) <= 0:
        await message.answer("⚠️ Введите число больше 0" if lang == "ru" else "⚠️ 0-ден үлкен санды енгізіңіз")
        return
    await state.update_data(tv_quantity=int(message.text))
    msg = (
        f"👨‍💼 <b>Шаг 6/8</b> — Сколько операторов/инструкторов?\n\n"
        f"💰 {OPERATOR_PRICE:,} тг/час\n<i>Введите 0, если не нужны</i>"
    ) if lang == "ru" else (
        f"👨‍💼 <b>Қадам 6/8</b> — Қанша оператор/нұсқаушы керек?\n\n"
        f"💰 {OPERATOR_PRICE:,} тг/сағ\n<i>Қажет болмаса 0 енгізіңіз</i>"
    )
    await message.answer(msg, parse_mode="HTML", reply_markup=cancel_menu())
    await state.set_state(RentForm.operators)

@dp.message(RentForm.operators)
async def rent_operators(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    if not message.text.isdigit() or int(message.text) < 0:
        await message.answer("⚠️ Введите число (0 или больше)" if lang == "ru" else "⚠️ Санды енгізіңіз (0 немесе одан көп)")
        return
    await state.update_data(operators=int(message.text))
    msg = (
        "🎬 <b>Шаг 7/8</b> — Медиа-сопровождение?\n\n"
        "Выберите пакет съёмки:"
    ) if lang == "ru" else (
        "🎬 <b>Қадам 7/8</b> — Медиа-сүйемелдеу?\n\n"
        "Түсіру пакетін таңдаңыз:"
    )
    await message.answer(msg, parse_mode="HTML", reply_markup=media_menu())
    await state.set_state(RentForm.media)

@dp.message(RentForm.media)
async def rent_media(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    valid_media = list(MEDIA_PACKAGES.keys())
    if message.text not in valid_media:
        await message.answer("⚠️ Выберите вариант из списка" if lang == "ru" else "⚠️ Тізімнен нұсқаны таңдаңыз")
        return
    media_choice = message.text
    media_cost = MEDIA_PACKAGES[media_choice]
    await state.update_data(media=media_choice, media_cost=media_cost)

    # Показываем детальный чек перед оформлением
    data = await state.get_data()
    equipment = data['equipment']
    quantity = data.get('quantity', 1)
    hours = data['hours']
    tv = data.get('tv', False)
    tv_qty = data.get('tv_quantity', 0)
    operators = data.get('operators', 0)

    price_per_hour = PRICE_BULK if quantity >= 5 else PRICE_STANDARD
    equipment_total = price_per_hour * quantity * hours
    tv_total = TV_PRICE * tv_qty * hours if tv else 0
    operator_total = OPERATOR_PRICE * operators * hours
    subtotal = equipment_total + tv_total + operator_total + media_cost + TRANSPORT_PRICE

    bulk_line = f"\n   🔥 <i>Оптовая цена 7 500 тг/час (5+ шлемов)</i>" if quantity >= 5 else ""
    tv_line = f"\n🔹 2. Дополнительное оборудование:\n   📺 Телевизоры — {tv_qty} шт. × {hours} ч. = <b>{tv_total:,} тг</b>" if tv else ""
    operator_line = f"\n🔹 3. Обслуживание и персонал:\n   👨‍💼 Операторы/инструкторы — {operators} чел. × {hours} ч. = <b>{operator_total:,} тг</b>" if operators > 0 else ""
    media_line = f"Выбранный пакет: {media_choice} = <b>от {media_cost:,} тг</b>" if media_cost > 0 else "Не выбрано — <b>0 тг</b>"

    check_text = (
        f"🧾 <b>ДЕТАЛЬНЫЙ РАСЧЁТ ВАШЕГО ЗАКАЗА:</b>\n\n"
        f"🔹 1. Основное VR-оборудование:\n"
        f"   🎮 {equipment} — {quantity} шт. × {hours} ч. = <b>{equipment_total:,} тг</b>{bulk_line}"
        f"{tv_line}"
        f"{operator_line}\n"
        f"🔹 4. Медиа-сопровождение (бэкстейдж):\n"
        f"   {media_line}\n\n"
        f"🔹 5. Логистика и обеспечение:\n"
        f"   🚗 Доставка, монтаж, настройка сети и выездной контроль — <b>{TRANSPORT_PRICE:,} тг</b>\n\n"
        f"──────────────────────\n"
        f"💰 <b>ИТОГО К ОПЛАТЕ: {subtotal:,} тг</b>\n\n"
        f"<i>Всё оборудование предоставляется в чистом, продезинфицированном виде, "
        f"полностью заряженным и готовым к бесперебойной работе.</i>\n"
        f"──────────────────────\n\n"
        f"👇 Если всё верно, нажмите <b>«Оформить заявку»</b> и укажите ваши контакты."
    )
    await state.update_data(subtotal=subtotal, price_per_hour=price_per_hour,
                             equipment_total=equipment_total, tv_total=tv_total,
                             operator_total=operator_total)
    await message.answer(check_text, parse_mode="HTML", reply_markup=confirm_order_menu())
    await state.set_state(RentForm.name)

@dp.message(RentForm.name)
async def rent_name(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    if message.text == "✅ Оформить заявку":
        msg = "👤 <b>Шаг 8/8</b> — Введите ваше имя:" if lang == "ru" else "👤 <b>Қадам 8/8</b> — Атыңызды енгізіңіз:"
        await message.answer(msg, parse_mode="HTML", reply_markup=cancel_menu())
        await state.update_data(confirmed=True)
        return
    data = await state.get_data()
    if not data.get('confirmed'):
        return
    await state.update_data(name=message.text)
    msg = "📱 Введите номер телефона:" if lang == "ru" else "📱 Телефон нөміріңізді енгізіңіз:"
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
    tv_qty = data.get('tv_quantity', 0)
    operators = data.get('operators', 0)
    media = data.get('media', '❌ Без медиа-сопровождения')
    media_cost = data.get('media_cost', 0)
    price_per_hour = data.get('price_per_hour', PRICE_STANDARD)
    equipment_total = data.get('equipment_total', 0)
    tv_total = data.get('tv_total', 0)
    operator_total = data.get('operator_total', 0)
    total = data.get('subtotal', 0)

    user = message.from_user

    order = {
        "equipment": equipment,
        "quantity": quantity,
        "hours": hours,
        "tv": tv,
        "tv_quantity": tv_qty,
        "operators": operators,
        "media": media,
        "media_cost": media_cost,
        "price_per_hour": price_per_hour,
        "equipment_cost": equipment_total,
        "tv_cost": tv_total,
        "operator_cost": operator_total,
        "transport_cost": TRANSPORT_PRICE,
        "total": total,
        "name": data.get('name', '—'),
        "phone": data.get('phone', '—'),
        "address": data.get('address', '—'),
        "comments": data.get('comments', '—'),
        "user_id": user.id,
        "username": user.username or "нет",
        "lang": lang,
        "promo": None,
        "discount": 0,
    }

    order_id = add_order(order)

    bulk_note = " 🔥 (оптовая цена)" if quantity >= 5 else ""
    admin_text = (
        f"🔥 <b>НОВАЯ ЗАЯВКА #{order_id}</b>\n"
        f"📅 {datetime.now().strftime('%d.%m.%Y %H:%M')}\n\n"
        f"🔹 1. VR-оборудование:\n"
        f"   {equipment} — {quantity} шт. × {hours} ч. × {price_per_hour:,} тг{bulk_note} = {equipment_total:,} тг\n"
    )
    if tv:
        admin_text += f"🔹 2. Телевизоры:\n   {tv_qty} шт. × {hours} ч. × {TV_PRICE:,} тг = {tv_total:,} тг\n"
    if operators > 0:
        admin_text += f"🔹 3. Персонал:\n   {operators} чел. × {hours} ч. × {OPERATOR_PRICE:,} тг = {operator_total:,} тг\n"
    admin_text += (
        f"🔹 4. Медиа: {media} = {media_cost:,} тг\n"
        f"🔹 5. Логистика: {TRANSPORT_PRICE:,} тг\n"
        f"─────────────────\n"
        f"💰 <b>ИТОГО: {total:,} тг</b>\n\n"
        f"👤 {data.get('name', '—')} | 📱 {data.get('phone', '—')}\n"
        f"📍 {data.get('address', '—')}\n"
        f"💬 {data.get('comments', '—')}\n"
        f"Telegram: @{user.username or 'нет'} | ID: <code>{user.id}</code>"
    )

    try:
        for admin_id in ADMIN_IDS:
            await bot.send_message(admin_id, admin_text, parse_mode="HTML", reply_markup=status_menu(order_id))
    except Exception as e:
        logger.error(f"Ошибка отправки заявки: {e}")

    confirm_text = (
        f"✅ <b>Заявка #{order_id} принята!</b>\n\n"
        f"🎮 {equipment} × {quantity} шт. × {hours} ч.\n"
        f"💰 <b>Итого: {total:,} тг</b>\n\n"
        f"Менеджер свяжется с вами в течение часа 🚀"
    ) if lang == "ru" else (
        f"✅ <b>Өтінім #{order_id} қабылданды!</b>\n\n"
        f"🎮 {equipment} × {quantity} дана × {hours} сағ.\n"
        f"💰 <b>Барлығы: {total:,} тг</b>\n\n"
        f"Менеджер бір сағат ішінде хабарласады 🚀"
    )

    await message.answer(confirm_text, parse_mode="HTML", reply_markup=main_menu(lang))
    await state.clear()


# ================== АРЕНДА КИНОКАМЕРЫ ==================
CAMERA_PRICE = 20000  # тг/час за кинокамеру

@dp.message(F.text.in_(["🎬 Аренда кинокамеры", "🎬 Кинокамера жалдау"]))
async def camera_start(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    if not is_working_hours():
        await message.answer(T["off_hours"][lang], parse_mode="HTML")
        return
    msg = (
        "🎬 <b>Аренда кинокамеры Sony FX30 Cinema Line</b>\n\n"
        "📸 Профессиональная съёмка вашего мероприятия\n\n"
        "<b>Шаг 1/4</b> — На сколько часов нужна кинокамера?\n"
        f"💰 от {CAMERA_PRICE:,} тг/час\n\n"
        "<i>Введите число часов (мин. 3)</i>"
    ) if lang == "ru" else (
        "🎬 <b>Sony FX30 Cinema Line кинокамерасын жалдау</b>\n\n"
        "📸 Іс-шараңызды кәсіби түсіру\n\n"
        "<b>Қадам 1/4</b> — Кинокамера қанша сағатқа керек?\n"
        f"💰 {CAMERA_PRICE:,} тг/сағ-дан\n\n"
        "<i>Сағат санын енгізіңіз (мин. 3)</i>"
    )
    await message.answer(msg, parse_mode="HTML", reply_markup=cancel_menu())
    await state.set_state(CameraForm.hours)

@dp.message(CameraForm.hours)
async def camera_hours(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    if not message.text.isdigit() or int(message.text) <= 0:
        await message.answer("⚠️ Введите количество часов числом" if lang == "ru" else "⚠️ Сағат санын енгізіңіз")
        return
    hours = int(message.text)
    if hours < MIN_HOURS:
        await message.answer(f"⚠️ Минимальный заказ — {MIN_HOURS} часа." if lang == "ru" else f"⚠️ Ең аз тапсырыс — {MIN_HOURS} сағат.")
        return
    camera_total = CAMERA_PRICE * hours
    total = camera_total + TRANSPORT_PRICE
    await state.update_data(hours=hours, camera_total=camera_total, total=total)

    check_text = (
        f"🧾 <b>РАСЧЁТ ЗАКАЗА:</b>\n\n"
        f"🎬 Sony FX30 Cinema Line — {hours} ч. = <b>{camera_total:,} тг</b>\n"
        f"🚗 Логистика (доставка + монтаж) — <b>{TRANSPORT_PRICE:,} тг</b>\n"
        f"──────────────────────\n"
        f"💰 <b>ИТОГО: {total:,} тг</b>\n\n"
        f"👇 Если всё верно, нажмите <b>«Оформить заявку»</b>"
    ) if lang == "ru" else (
        f"🧾 <b>ТАПСЫРЫС ЕСЕБІ:</b>\n\n"
        f"🎬 Sony FX30 Cinema Line — {hours} сағ. = <b>{camera_total:,} тг</b>\n"
        f"🚗 Логистика — <b>{TRANSPORT_PRICE:,} тг</b>\n"
        f"──────────────────────\n"
        f"💰 <b>БАРЛЫҒЫ: {total:,} тг</b>\n\n"
        f"👇 Барлығы дұрыс болса, <b>«Өтінімді рәсімдеу»</b> батырмасын басыңыз"
    )
    await message.answer(check_text, parse_mode="HTML", reply_markup=confirm_order_menu())
    await state.set_state(CameraForm.name)

@dp.message(CameraForm.name)
async def camera_name(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    if message.text == "✅ Оформить заявку":
        msg = "👤 Введите ваше имя:" if lang == "ru" else "👤 Атыңызды енгізіңіз:"
        await message.answer(msg, reply_markup=cancel_menu())
        await state.update_data(confirmed=True)
        return
    data = await state.get_data()
    if not data.get("confirmed"):
        return
    await state.update_data(name=message.text)
    msg = "📱 Введите номер телефона:" if lang == "ru" else "📱 Телефон нөміріңізді енгізіңіз:"
    await message.answer(msg, reply_markup=cancel_menu())
    await state.set_state(CameraForm.phone)

@dp.message(CameraForm.phone)
async def camera_phone(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    await state.update_data(phone=message.text)
    msg = "📍 Укажите адрес съёмки (Алматы):" if lang == "ru" else "📍 Түсіру мекенжайын көрсетіңіз (Алматы):"
    await message.answer(msg, reply_markup=cancel_menu())
    await state.set_state(CameraForm.address)

@dp.message(CameraForm.address)
async def camera_address(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    await state.update_data(address=message.text)
    msg = "💬 Дополнительные пожелания? (или напишите «нет»)" if lang == "ru" else "💬 Қосымша тілектер? (немесе «жоқ» деп жазыңыз)"
    await message.answer(msg, reply_markup=cancel_menu())
    await state.set_state(CameraForm.comments)

@dp.message(CameraForm.comments)
async def camera_finish(message: types.Message, state: FSMContext):
    lang = get_user_lang(message.from_user.id)
    await state.update_data(comments=message.text)
    data = await state.get_data()

    hours = data["hours"]
    camera_total = data["camera_total"]
    total = data["total"]
    user = message.from_user

    order = {
        "type": "camera",
        "equipment": "Sony FX30 Cinema Line",
        "hours": hours,
        "camera_cost": camera_total,
        "transport_cost": TRANSPORT_PRICE,
        "total": total,
        "name": data.get("name", "—"),
        "phone": data.get("phone", "—"),
        "address": data.get("address", "—"),
        "comments": data.get("comments", "—"),
        "user_id": user.id,
        "username": user.username or "нет",
        "lang": lang,
        "promo": None,
        "discount": 0,
    }

    order_id = add_order(order)

    admin_text = (
        f"🎬 <b>ЗАЯВКА НА КИНОКАМЕРУ #{order_id}</b>\n"
        f"📅 {datetime.now().strftime('%d.%m.%Y %H:%M')}\n\n"
        f"📷 Sony FX30 Cinema Line — {hours} ч. = {camera_total:,} тг\n"
        f"🚗 Логистика: {TRANSPORT_PRICE:,} тг\n"
        f"─────────────────\n"
        f"💰 <b>ИТОГО: {total:,} тг</b>\n\n"
        f"👤 {data.get('name', '—')} | 📱 {data.get('phone', '—')}\n"
        f"📍 {data.get('address', '—')}\n"
        f"💬 {data.get('comments', '—')}\n"
        f"Telegram: @{user.username or 'нет'} | ID: <code>{user.id}</code>"
    )

    try:
        for admin_id in ADMIN_IDS:
            await bot.send_message(admin_id, admin_text, parse_mode="HTML", reply_markup=status_menu(order_id))
    except Exception as e:
        logger.error(f"Ошибка отправки заявки кинокамеры: {e}")

    confirm_text = (
        f"✅ <b>Заявка #{order_id} принята!</b>\n\n"
        f"🎬 Sony FX30 Cinema Line — {hours} ч.\n"
        f"💰 <b>Итого: {total:,} тг</b>\n\n"
        f"Менеджер свяжется с вами в течение часа 🚀"
    ) if lang == "ru" else (
        f"✅ <b>Өтінім #{order_id} қабылданды!</b>\n\n"
        f"🎬 Sony FX30 Cinema Line — {hours} сағ.\n"
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

    headers = ["#", "Дата", "Статус", "Оборудование", "Шлемов", "Часов", "ТВ", "Кол-во ТВ",
               "Операторов", "Медиа", "Цена/час", "Шлемы сумма", "ТВ сумма",
               "Персонал", "Медиа сумма", "Транспорт", "Итого",
               "Имя", "Телефон", "Адрес", "Комментарий", "Telegram"]
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
        ws.cell(row=row, column=8, value=o.get("tv_quantity", 0))
        ws.cell(row=row, column=9, value=o.get("operators", 0))
        ws.cell(row=row, column=10, value=o.get("media", "—"))
        ws.cell(row=row, column=11, value=o.get("price_per_hour", 0))
        ws.cell(row=row, column=12, value=o.get("equipment_cost", 0))
        ws.cell(row=row, column=13, value=o.get("tv_cost", 0))
        ws.cell(row=row, column=14, value=o.get("operator_cost", 0))
        ws.cell(row=row, column=15, value=o.get("media_cost", 0))
        ws.cell(row=row, column=16, value=o.get("transport_cost", 0))
        ws.cell(row=row, column=17, value=o.get("total", 0))
        ws.cell(row=row, column=18, value=o.get("name"))
        ws.cell(row=row, column=19, value=o.get("phone"))
        ws.cell(row=row, column=20, value=o.get("address"))
        ws.cell(row=row, column=21, value=o.get("comments"))
        ws.cell(row=row, column=22, value=f"@{o.get('username', 'нет')}")

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
        f"🎮 {o.get('equipment')} x{o.get('quantity', 1)} | {o.get('hours', '—')} ч.\n"
        f"📺 ТВ: {'Да (' + str(o.get('tv_quantity', 1)) + ' шт.)' if o.get('tv') else 'Нет'}\n"
        f"👨‍💼 Операторов: {o.get('operators', 0)}\n"
        f"🎬 Медиа: {o.get('media', '—')}\n"
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

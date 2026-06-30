"""
Bestpresso Tech Bot — бот для техпідтримки (диспетчер/технік створює заявку,
технік закриває заявку текстовим повідомленням з кодом у групі).

Змінні середовища (Render або .env):
  TECHBOT_TOKEN       — токен бота від @BotFather (НОВИЙ бот, окремий від helpbot)
  GOOGLE_CREDENTIALS  — вміст credentials.json (той самий сервісний акаунт, що й в helpbot)

Група для заявок техпідтримки: ЗАМІНІТЬ на свій chat_id нижче (TECH_CHAT_ID)
Таблиця: НОВА Google Таблиця — створіть і вставте її ID нижче (TECH_SPREADSHEET_ID)

──────────────────────────────────────────────────────────────────────────────
ЯК НАЛАШТУВАТИ:
1. Створіть нового бота через @BotFather → отримайте токен → TECHBOT_TOKEN.
2. Створіть нову Google Таблицю, на першому аркуші назвіть колонки як у SHEET_NAME
   нижче (бот сам допише шапку, якщо аркуш порожній).
3. Дайте доступ сервісному акаунту (той самий email з credentials.json) до нової
   таблиці як Редактор, скопіюйте ID таблиці з URL → TECH_SPREADSHEET_ID.
4. Додайте бота в групу техпідтримки, дізнайтесь chat_id групи (через getUpdates
   або @userinfobot) → TECH_CHAT_ID.
5. Відредагуйте список LOCATIONS нижче — впишіть свої 8 локацій.
6. Запустіть: python techbot.py
──────────────────────────────────────────────────────────────────────────────
"""

import os
import re
import json
import logging
from datetime import datetime, timezone, timedelta, time as dtime

import gspread
from google.oauth2.service_account import Credentials

from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
    KeyboardButton,
    ReactionTypeEmoji,
)
from telegram.ext import (
    Application,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    filters,
    ContextTypes,
)

# ── Логування ─────────────────────────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# ── Константи (ЗАМІНІТЬ під себе) ───────────────────────────────────────────────
BOT_TOKEN          = os.getenv("TECHBOT_TOKEN", "")
TECH_SPREADSHEET_ID = "1i9XhPVGnPPlxSyYBlgu5a_lIScRoGB-FPPqy0doOzow"  # ID нової Google Таблиці
SHEET_NAME          = "Заявки"                              # назва аркуша
TECH_CHAT_ID         = -5565166269                           # chat_id групи техпідтримки
KYIV_TZ              = timezone(timedelta(hours=3))

# 8 локацій
LOCATIONS = [
    "Конєва",
    "Метрологічна",
    "Зоо №1 (більша будка)",
    "Андріївський",
    "Малевича",
    "Протасів Яр",
    "Предславинська",
    "Зоо №2 (менша будка)",
]

# Шапка таблиці (якщо аркуш порожній — бот допише сам)
HEADER = [
    "Код заявки", "Дата заявки", "Час заявки", "Локація", "Номер телефону",
    "Проблематика", "Фото / відео", "Що зроблено", "Виконавець",
    "Дата закриття", "Час закриття",
]

# Літери колонок у таблиці (A=1 ... K=11)
COL_CODE, COL_DATE, COL_TIME, COL_LOCATION, COL_PHONE, \
    COL_PROBLEM, COL_MEDIA, COL_WORK, COL_EXECUTOR, COL_CLOSED_DATE, COL_CLOSED_TIME = range(1, 12)

DEFAULT_EXECUTOR = "Боковенко"          # виконавець 1 — проставляється одразу при створенні
EXECUTOR_GREEN = {"red": 0.0, "green": 1.0, "blue": 0.0}  # #00ff00

CODE_PATTERN = re.compile(r"^#?(\d{1,6})\b\s*(.*)", re.DOTALL)

WELCOME_TEXT = "👋Вас вітає бот Техпідтримки. Опишіть проблему, яка виникла на локації."
ASK_PROBLEM_TEXT = "Опишіть проблему, яка виникла на локації."
ASK_PHONE_TEXT = "Поділіться, будь ласка, номером телефону для зв'язку (потрібно лише один раз):"
CONFIRM_TEXT = "✅ Вашу заявку прийнято та передано в Техпідтримку."

USERS_SHEET_NAME = "Користувачі"
USERS_HEADER = ["Telegram ID", "Номер телефону", "Дата збереження"]
LIKE_REACTION = "👍"



# ── Google Sheets ─────────────────────────────────────────────────────────────

def get_gspread_client():
    """Повертає авторизований gspread клієнт (той самий принцип, що в helpbot.py)."""
    scopes = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds_file = os.path.join(os.path.dirname(__file__), "credentials.json")
    if os.path.exists(creds_file):
        creds = Credentials.from_service_account_file(creds_file, scopes=scopes)
        return gspread.authorize(creds)

    creds_json = os.getenv("GOOGLE_CREDENTIALS", "").strip()
    if not creds_json:
        raise ValueError("Не задано GOOGLE_CREDENTIALS і немає credentials.json!")

    try:
        creds_dict = json.loads(creds_json)
    except json.JSONDecodeError:
        if creds_json.startswith('"') and creds_json.endswith('"'):
            creds_json = creds_json[1:-1]
        creds_json = creds_json.encode("utf-8").decode("unicode_escape")
        creds_dict = json.loads(creds_json)

    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)


def get_sheet():
    client = get_gspread_client()
    spreadsheet = client.open_by_key(TECH_SPREADSHEET_ID)
    try:
        sheet = spreadsheet.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title=SHEET_NAME, rows=2000, cols=len(HEADER))
        sheet.append_row(HEADER)
        return sheet

    # Якщо аркуш існує, але порожній — допишемо шапку
    if not sheet.get_all_values():
        sheet.append_row(HEADER)
    return sheet


def get_users_sheet():
    """Окрема вкладка 'Користувачі' — зберігає номери телефонів по chat_id,
    щоб вони не губились при перезапуску бота на Render."""
    client = get_gspread_client()
    spreadsheet = client.open_by_key(TECH_SPREADSHEET_ID)
    try:
        sheet = spreadsheet.worksheet(USERS_SHEET_NAME)
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title=USERS_SHEET_NAME, rows=1000, cols=len(USERS_HEADER))
        sheet.append_row(USERS_HEADER)
        return sheet

    if not sheet.get_all_values():
        sheet.append_row(USERS_HEADER)
    return sheet


def load_phone_cache() -> dict:
    """Зчитує всю вкладку 'Користувачі' одним запитом і повертає словник
    {telegram_id (str): phone}. Викликається один раз при старті бота."""
    cache = {}
    try:
        sheet = get_users_sheet()
        rows = sheet.get_all_values()[1:]  # пропускаємо шапку
        for row in rows:
            if len(row) >= 2 and row[0].strip():
                cache[row[0].strip()] = row[1].strip()
        logger.info(f"Завантажено {len(cache)} збережених номерів телефонів")
    except Exception as e:
        logger.error(f"Не вдалося завантажити кеш номерів телефонів: {e}")
    return cache


def save_user_phone(chat_id: int, phone: str):
    """Дописує новий рядок у вкладку 'Користувачі'."""
    sheet = get_users_sheet()
    now = datetime.now(KYIV_TZ).strftime("%d.%m.%Y %H:%M:%S")
    sheet.append_row([str(chat_id), phone, now], value_input_option="USER_ENTERED")


def next_ticket_number(sheet) -> int:
    """Шукає максимальний номер коду в колонці A і повертає наступний."""
    col_a = sheet.col_values(COL_CODE)[1:]  # пропускаємо шапку
    last = 0
    for val in col_a:
        val = val.strip().lstrip("#")
        if val.isdigit():
            last = max(last, int(val))
    return last + 1


def append_ticket(code, location, phone, problem, has_media):
    """Додає новий рядок заявки. Виконавець проставляється одразу (зелений колір)."""
    sheet = get_sheet()
    now = datetime.now(KYIV_TZ)
    date_str = now.strftime("%d.%m.%Y")
    time_str = now.strftime("%H:%M:%S")
    code_str = f"{code:04d}"
    row = [
        code_str, date_str, time_str, location, phone,
        problem, "є в телеграмі" if has_media else "немає",
        "", DEFAULT_EXECUTOR, "", "",
    ]
    sheet.append_row(row, value_input_option="USER_ENTERED")
    row_num = len(sheet.col_values(COL_CODE))

    # Заливаємо клітинку "Виконавець" зеленим (#00ff00)
    try:
        sheet.format(f"I{row_num}", {"backgroundColor": EXECUTOR_GREEN})
    except Exception as e:
        logger.warning(f"Не вдалося пофарбувати клітинку виконавця: {e}")

    return code_str


def close_ticket(code: str, work_description: str) -> bool:
    """Знаходить рядок за кодом заявки і записує виконані роботи + дату/час закриття.
    Зелена заливка клітинки виконавця лишається незмінною (і для відкритих, і для закритих)."""
    sheet = get_sheet()
    col_a = sheet.col_values(COL_CODE)
    target_row = None
    code_norm = code.lstrip("#").strip().lstrip("0") or "0"
    for i, val in enumerate(col_a, start=1):
        val_norm = val.strip().lstrip("#").lstrip("0") or "0"
        if val_norm == code_norm:
            target_row = i
            break

    if not target_row:
        return False

    now = datetime.now(KYIV_TZ)
    updates = [
        {"range": f"H{target_row}", "values": [[work_description]]},
        {"range": f"J{target_row}", "values": [[now.strftime("%d.%m.%Y")]]},
        {"range": f"K{target_row}", "values": [[now.strftime("%H:%M:%S")]]},
    ]
    sheet.batch_update(updates, value_input_option="USER_ENTERED")
    return True


def get_unclosed_tickets() -> list:
    """Повертає список незакритих заявок — тих, де порожня колонка 'Дата закриття'
    (зелена заливка в колонці 'Виконавець' лишається завжди, незалежно від статусу)."""
    sheet = get_sheet()
    rows = sheet.get_all_values()[1:]  # пропускаємо шапку
    unclosed = []
    for row in rows:
        row = (row + [""] * 11)[:11]
        code, date_, time_, location, phone, problem, media, work, executor, closed_date, closed_time = row
        if not code.strip():
            continue
        if not closed_date.strip():
            unclosed.append({
                "code": code.strip(),
                "date": date_.strip(),
                "time": time_.strip(),
                "location": location.strip(),
                "phone": phone.strip(),
                "problem": problem.strip(),
            })
    return unclosed


def format_unclosed_report(tickets: list) -> str:
    if not tickets:
        return "✅ Усі заявки закриті. Незакритих немає."

    lines = [f"📋 *Незакриті заявки ({len(tickets)}):*\n"]
    for t in tickets:
        lines.append(
            f"🔸 *{t['code']}* | {t['location']} | {t['date']} {t['time']}\n"
            f"{t['problem']}\n"
            f"📞 {t['phone']}"
        )
    return "\n\n".join(lines)


# ── Клавіатури ────────────────────────────────────────────────────────────────

def locations_keyboard():
    buttons = [
        [InlineKeyboardButton(loc, callback_data=f"loc_{i}")]
        for i, loc in enumerate(LOCATIONS)
    ]
    return InlineKeyboardMarkup(buttons)


def skip_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("⏭ Пропустити", callback_data="skip_media")]
    ])


def phone_request_keyboard():
    return ReplyKeyboardMarkup(
        [[KeyboardButton("📞 Поділитися номером", request_contact=True)]],
        resize_keyboard=True,
        one_time_keyboard=True,
    )


def new_ticket_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🆕 Нова заявка", callback_data="new_ticket")]
    ])


# ── Ручна машина станів (без ConversationHandler — надійніше) ───────────────
# context.user_data["step"]: None -> ["await_phone" якщо новий] -> "problem" -> "location" -> "media"

def get_cached_phone(context: ContextTypes.DEFAULT_TYPE, chat_id: int):
    return context.bot_data.setdefault("phones", {}).get(str(chat_id))


async def handle_private_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Єдиний обробник тексту в приватному чаті. Реагує на будь-яке повідомлення,
    /start не обов'язковий."""
    step = context.user_data.get("step")
    chat_id = update.effective_chat.id

    # Немає активного діалогу — починаємо новий
    if step is None:
        context.user_data.clear()
        phone = get_cached_phone(context, chat_id)
        if phone:
            context.user_data["phone"] = phone
            context.user_data["step"] = "problem"
            await update.message.reply_text(WELCOME_TEXT)
        else:
            context.user_data["step"] = "await_phone"
            await update.message.reply_text(WELCOME_TEXT)
            await update.message.reply_text(
                ASK_PHONE_TEXT,
                reply_markup=phone_request_keyboard(),
            )
        return

    if step == "await_phone":
        await update.message.reply_text(
            ASK_PHONE_TEXT,
            reply_markup=phone_request_keyboard(),
        )
        return

    if step == "problem":
        context.user_data["problem"] = update.message.text.strip()
        context.user_data["step"] = "location"
        await update.message.reply_text(
            "Оберіть локацію:",
            reply_markup=locations_keyboard(),
        )
        return

    if step == "location":
        await update.message.reply_text(
            "Будь ласка, оберіть локацію кнопкою вище ⬆️",
            reply_markup=locations_keyboard(),
        )
        return

    if step == "media":
        await update.message.reply_text(
            "Додайте фото/відео або натисніть «Пропустити» нижче ⬇️",
            reply_markup=skip_keyboard(),
        )
        return


async def handle_private_contact(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Отримання номера телефону через кнопку 'Поділитися' (лише при першому зверненні)."""
    if context.user_data.get("step") != "await_phone":
        return

    contact = update.message.contact
    if not contact:
        return

    phone = contact.phone_number
    chat_id = update.effective_chat.id
    context.user_data["phone"] = phone

    try:
        save_user_phone(chat_id, phone)
        context.bot_data.setdefault("phones", {})[str(chat_id)] = phone
        logger.info(f"Збережено номер телефону для chat_id={chat_id}")
    except Exception as e:
        logger.error(f"Не вдалося зберегти номер телефону: {e}")

    context.user_data["step"] = "problem"
    await update.message.reply_text(
        "Дякуємо! Номер збережено.",
        reply_markup=ReplyKeyboardRemove(),
    )
    await update.message.reply_text(ASK_PROBLEM_TEXT)


async def handle_private_media(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Фото/відео в приваті — приймається лише на кроці 'media'."""
    if context.user_data.get("step") != "media":
        return

    media_type = None
    media_id = None
    if update.message.photo:
        media_type = "photo"
        media_id = update.message.photo[-1].file_id
    elif update.message.video:
        media_type = "video"
        media_id = update.message.video.file_id

    context.user_data["media_type"] = media_type
    context.user_data["media_id"] = media_id
    await finalize_ticket(update, context)


async def handle_location_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if context.user_data.get("step") != "location":
        return  # натиснута стара кнопка з минулої заявки — ігноруємо

    idx = int(query.data.split("_")[1])
    context.user_data["location"] = LOCATIONS[idx]
    context.user_data["step"] = "media"

    await query.edit_message_text(
        f"📍 Локація: *{LOCATIONS[idx]}*",
        parse_mode="Markdown",
    )
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text="📷 Додайте фото/відео (якщо є) або натисніть кнопку нижче:",
        reply_markup=skip_keyboard(),
    )


async def handle_skip_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if context.user_data.get("step") != "media":
        return

    context.user_data["media_type"] = None
    context.user_data["media_id"] = None
    await finalize_ticket(update, context, from_callback=True)


async def handle_new_ticket_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Кнопка 'Нова заявка' після підтвердження — одразу питає про проблему,
    минаючи вітання (телефон уже відомий)."""
    query = update.callback_query
    await query.answer()

    chat_id = update.effective_chat.id
    phone = get_cached_phone(context, chat_id)
    context.user_data.clear()
    context.user_data["phone"] = phone
    context.user_data["step"] = "problem"

    await context.bot.send_message(chat_id=chat_id, text=ASK_PROBLEM_TEXT)


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/start — примусово скидає й починає заявку заново."""
    chat_id = update.effective_chat.id
    phone = get_cached_phone(context, chat_id)
    context.user_data.clear()

    if phone:
        context.user_data["phone"] = phone
        context.user_data["step"] = "problem"
        await update.message.reply_text(WELCOME_TEXT)
    else:
        context.user_data["step"] = "await_phone"
        await update.message.reply_text(WELCOME_TEXT)
        await update.message.reply_text(
            ASK_PHONE_TEXT,
            reply_markup=phone_request_keyboard(),
        )


# ── Фіналізація заявки: відправка в групу + запис у таблицю ─────────────────

async def finalize_ticket(update: Update, context: ContextTypes.DEFAULT_TYPE, from_callback=False):
    location = context.user_data.get("location", "—")
    problem = context.user_data.get("problem", "—")
    phone = context.user_data.get("phone", "—")
    media_type = context.user_data.get("media_type")
    media_id = context.user_data.get("media_id")

    sheet = get_sheet()
    next_num = next_ticket_number(sheet)

    try:
        code_str = append_ticket(next_num, location, phone, problem, bool(media_type))
        logger.info(f"Заявку #{code_str} записано в таблицю")
    except Exception as e:
        logger.error(f"Помилка запису заявки в таблицю: {e}")
        code_str = f"{next_num:04d}"

    caption = (
        f"🆕 *Нова заявка #{code_str}*\n\n"
        f"📍 Локація: {location}\n"
        f"📝 Проблема: {problem}\n"
        f"📞 Телефон: {phone}\n\n"
        f"_Щоб закрити — напишіть у групі: {code_str} і опис виконаних робіт_"
    )

    bot = context.bot
    try:
        if media_type == "photo" and media_id:
            await bot.send_photo(chat_id=TECH_CHAT_ID, photo=media_id, caption=caption, parse_mode="Markdown")
        elif media_type == "video" and media_id:
            await bot.send_video(chat_id=TECH_CHAT_ID, video=media_id, caption=caption, parse_mode="Markdown")
        else:
            await bot.send_message(chat_id=TECH_CHAT_ID, text=caption, parse_mode="Markdown")
        logger.info(f"Заявку #{code_str} надіслано в групу техпідтримки")
    except Exception as e:
        logger.error(f"Помилка надсилання заявки в групу: {e}")

    if from_callback:
        await update.callback_query.edit_message_text(CONFIRM_TEXT, reply_markup=new_ticket_keyboard())
    else:
        await update.message.reply_text(CONFIRM_TEXT, reply_markup=new_ticket_keyboard())

    context.user_data.clear()


# ── Звіт по незакритих заявках ───────────────────────────────────────────────

async def send_unclosed_report(context: ContextTypes.DEFAULT_TYPE, chat_id: int = None):
    """Формує і надсилає список незакритих заявок. За замовчуванням — у групу техпідтримки."""
    target_chat = chat_id or TECH_CHAT_ID
    try:
        tickets = get_unclosed_tickets()
        text = format_unclosed_report(tickets)
    except Exception as e:
        logger.error(f"Помилка формування звіту: {e}")
        text = "⚠️ Не вдалося сформувати звіт — помилка доступу до таблиці."

    await context.bot.send_message(chat_id=target_chat, text=text, parse_mode="Markdown")


async def daily_unclosed_report_job(context: ContextTypes.DEFAULT_TYPE):
    """Автоматична розсилка о 17:00 (пн-пт + нд) — викликається JobQueue."""
    logger.info("Запуск щоденного звіту незакритих заявок")
    await send_unclosed_report(context)


# ── Закриття заявки в групі (звичайне текстове повідомлення, без Reply) ──────

async def handle_group_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Технік у групі техпідтримки пише: '0001 замінив датчик, перевірено'
    Бот парсить код на початку повідомлення і закриває заявку.
    Команда '/звіт' (текстом, кирилицею) — надсилає список незакритих заявок.
    """
    message = update.message
    if not message or message.chat_id != TECH_CHAT_ID or not message.text:
        return

    text_stripped = message.text.strip()

    # Команда /звіт (Telegram не розпізнає кириличні команди як bot_command,
    # тому перевіряємо текст напряму)
    if text_stripped.lower().lstrip("/") == "звіт":
        await send_unclosed_report(context, chat_id=message.chat_id)
        return

    match = CODE_PATTERN.match(text_stripped)
    if not match:
        return  # повідомлення не починається з коду — ігноруємо

    code, work_description = match.groups()
    work_description = work_description.strip()
    if not work_description:
        await message.reply_text(
            f"⚠️ Додайте опис виконаних робіт після коду, наприклад:\n{code} замінив датчик"
        )
        return

    try:
        ok = close_ticket(code, work_description)
    except Exception as e:
        logger.error(f"Помилка закриття заявки {code}: {e}")
        ok = False

    if ok:
        try:
            await context.bot.set_message_reaction(
                chat_id=message.chat_id,
                message_id=message.message_id,
                reaction=[ReactionTypeEmoji(LIKE_REACTION)],
            )
        except Exception as e:
            logger.error(f"Не вдалося поставити реакцію на повідомлення: {e}")
        logger.info(f"Заявку {code} закрито")
    else:
        await message.reply_text(f"⚠️ Заявку {code} не знайдено в таблиці. Перевірте номер.")


# ── Запуск ────────────────────────────────────────────────────────────────────

def main():
    if not BOT_TOKEN:
        raise ValueError("Не задано TECHBOT_TOKEN у змінних середовища!")

    import requests
    try:
        r = requests.post(
            f"https://api.telegram.org/bot{BOT_TOKEN}/deleteWebhook",
            json={"drop_pending_updates": True},
            timeout=10,
        )
        logger.info(f"Webhook cleanup: {r.json()}")
    except Exception as e:
        logger.warning(f"Webhook cleanup failed: {e}")

    app = Application.builder().token(BOT_TOKEN).build()

    # Завантажуємо збережені номери телефонів у пам'ять (переживає рестарти,
    # бо джерело — вкладка "Користувачі" в Google Таблиці)
    try:
        app.bot_data["phones"] = load_phone_cache()
    except Exception as e:
        logger.error(f"Помилка ініціалізації кешу телефонів: {e}")
        app.bot_data["phones"] = {}

    # Приватний чат: текст, медіа, контакт, кнопки локації/пропуску/нова заявка
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(MessageHandler(filters.ChatType.PRIVATE & filters.CONTACT, handle_private_contact))
    app.add_handler(MessageHandler(
        filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND,
        handle_private_text,
    ))
    app.add_handler(MessageHandler(
        filters.ChatType.PRIVATE & (filters.PHOTO | filters.VIDEO),
        handle_private_media,
    ))
    app.add_handler(CallbackQueryHandler(handle_location_callback, pattern=r"^loc_\d+$"))
    app.add_handler(CallbackQueryHandler(handle_skip_callback, pattern="^skip_media$"))
    app.add_handler(CallbackQueryHandler(handle_new_ticket_callback, pattern="^new_ticket$"))

    # Звичайні текстові повідомлення в групі техпідтримки — обробка закриття заявок + /звіт
    app.add_handler(
        MessageHandler(
            filters.TEXT & ~filters.COMMAND & filters.Chat(chat_id=TECH_CHAT_ID),
            handle_group_message,
        )
    )

    # Щоденний автозвіт о 17:00 (Київ): пн, вт, ср, чт, пт, нд (без суботи)
    # 0=понеділок ... 6=неділя
    app.job_queue.run_daily(
        daily_unclosed_report_job,
        time=dtime(hour=17, minute=0, tzinfo=KYIV_TZ),
        days=(0, 1, 2, 3, 4, 6),
        name="daily_unclosed_report",
    )

    logger.info("Bestpresso Tech Bot запущено ✅")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()

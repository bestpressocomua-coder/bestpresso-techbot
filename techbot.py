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
from datetime import datetime, timezone, timedelta

import gspread
from google.oauth2.service_account import Credentials

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    ConversationHandler,
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
CONFIRM_TEXT = "✅ Вашу заявку прийнято та передано в Техпідтримку."

# ── Стани ConversationHandler ────────────────────────────────────────────────
WAIT_PROBLEM, WAIT_LOCATION, WAIT_PHONE, WAIT_MEDIA = range(1, 5)

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
    """Знаходить рядок за кодом заявки і записує виконані роботи + дату/час закриття."""
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


# ── Старт розмови (будь-яке повідомлення у приваті, без /start) ─────────────

async def start_conversation(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(WELCOME_TEXT)
    return WAIT_PROBLEM


async def receive_problem(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["problem"] = update.message.text.strip()
    await update.message.reply_text(
        "Оберіть локацію:",
        reply_markup=locations_keyboard(),
    )
    return WAIT_LOCATION


async def receive_location(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    idx = int(query.data.split("_")[1])
    context.user_data["location"] = LOCATIONS[idx]

    await query.edit_message_text(
        f"📍 Локація: *{LOCATIONS[idx]}*\n\nВкажіть номер телефону для зв'язку:",
        parse_mode="Markdown",
    )
    return WAIT_PHONE


async def receive_phone(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.contact:
        phone = update.message.contact.phone_number
    else:
        phone = update.message.text.strip()
    context.user_data["phone"] = phone

    await update.message.reply_text(
        "📷 Додайте фото/відео (якщо є) або натисніть кнопку нижче:",
        reply_markup=skip_keyboard(),
    )
    return WAIT_MEDIA


async def receive_media(update: Update, context: ContextTypes.DEFAULT_TYPE):
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
    return await finalize_ticket(update, context)


async def skip_media(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    context.user_data["media_type"] = None
    context.user_data["media_id"] = None
    return await finalize_ticket(update, context, from_callback=True)


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
        await update.callback_query.edit_message_text(CONFIRM_TEXT)
    else:
        await update.message.reply_text(CONFIRM_TEXT)

    context.user_data.clear()
    return ConversationHandler.END


# ── Закриття заявки в групі (звичайне текстове повідомлення, без Reply) ──────

async def handle_group_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Технік у групі техпідтримки пише: '0001 замінив датчик, перевірено'
    Бот парсить код на початку повідомлення і закриває заявку.
    """
    message = update.message
    if not message or message.chat_id != TECH_CHAT_ID or not message.text:
        return

    match = CODE_PATTERN.match(message.text.strip())
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
        await message.reply_text(f"✅ Заявку {code} закрито. Дякуємо!")
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

    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start", start_conversation),
            MessageHandler(filters.TEXT & ~filters.COMMAND & filters.ChatType.PRIVATE, start_conversation),
        ],
        states={
            WAIT_PROBLEM: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_problem),
            ],
            WAIT_LOCATION: [
                CallbackQueryHandler(receive_location, pattern=r"^loc_\d+$"),
            ],
            WAIT_PHONE: [
                MessageHandler((filters.TEXT & ~filters.COMMAND) | filters.CONTACT, receive_phone),
            ],
            WAIT_MEDIA: [
                MessageHandler(filters.PHOTO | filters.VIDEO, receive_media),
                CallbackQueryHandler(skip_media, pattern="^skip_media$"),
            ],
        },
        fallbacks=[CommandHandler("start", start_conversation)],
        per_user=True,
        per_chat=True,
        allow_reentry=True,
    )

    app.add_handler(conv)
    # Звичайні текстові повідомлення в групі техпідтримки — обробка закриття заявок
    app.add_handler(
        MessageHandler(
            filters.TEXT & ~filters.COMMAND & filters.Chat(chat_id=TECH_CHAT_ID),
            handle_group_message,
        )
    )

    logger.info("Bestpresso Tech Bot запущено ✅")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()

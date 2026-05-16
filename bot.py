"""
Moliya boti — Excel faylga yozuvchi Telegram bot
O'zbek tilida | v3.0
Faqat BOT_TOKEN kerak, hech qanday Google yo'q!
"""

import os
import logging
from datetime import datetime, date, timedelta
from pathlib import Path

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler,
    ConversationHandler, filters, ContextTypes
)
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ─── Sozlamalar ───────────────────────────────────────────────────────────────
BOT_TOKEN = "5072885311:AAFL68nhof38RVMI9CnBYn5Yxqc8RJlC33s"
EXCEL_DIR = Path("moliya_data")
EXCEL_DIR.mkdir(exist_ok=True)

MONTH_NAMES = ["", "Yanvar", "Fevral", "Mart", "Aprel", "May", "Iyun",
               "Iyul", "Avgust", "Sentabr", "Oktabr", "Noyabr", "Dekabr"]

CHIQIM_KATEGORIYALAR = [
    "🍔 Oziq-ovqat", "🚗 Transport", "🏥 Tibbiy/Dorixona",
    "📚 Ta'lim", "🏠 Uy/Kommunal", "📱 Telefon/Internet",
    "👗 Kiyim-kechak", "💼 Biznes", "🎭 Ko'ngil ochar",
    "💸 P2P o'tkazma", "🏦 Bank to'lovi", "🛒 Savdo marketi",
    "👤 Shaxsiy to'lov", "🔧 Texnika", "📦 Boshqa",
]
KIRIM_KATEGORIYALAR = [
    "💰 Maosh/Daromad", "🏪 Biznes tushumi", "💳 Karta to'ldirish",
    "🏦 Bank o'tkazma", "🎁 Sovg'a", "💹 Cashback", "📦 Boshqa kirim",
]

(SELECTING_CATEGORY, ENTERING_AMOUNT, ENTERING_DESCRIPTION,
 SELECTING_DATE, CONFIRMING) = range(5)

logging.basicConfig(format="%(asctime)s %(levelname)s %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)


# ─── Excel funksiyalar ────────────────────────────────────────────────────────
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", start_color="1F4E79")
HDR_FONT = Font(bold=True, color="FFFFFF", name="Arial", size=10)


def get_excel_path(year: int, month: int) -> Path:
    return EXCEL_DIR / f"{year}-{month:02d}-moliya.xlsx"


def get_or_create_workbook(year: int, month: int):
    path = get_excel_path(year, month)
    if path.exists():
        wb = openpyxl.load_workbook(path)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Tranzaksiyalar"
        headers = ["Sana", "Vaqt", "Tur", "Kategoriya", "Summa (UZS)", "Izoh"]
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.font = HDR_FONT
            cell.fill = HDR_FILL
            cell.alignment = Alignment(horizontal="center")
            cell.border = BORDER
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 8
        ws.column_dimensions["C"].width = 12
        ws.column_dimensions["D"].width = 22
        ws.column_dimensions["E"].width = 18
        ws.column_dimensions["F"].width = 30
        ws.freeze_panes = "A2"
        wb.save(path)
    return wb, path


def add_transaction(data: dict) -> str:
    now = datetime.now()
    wb, path = get_or_create_workbook(now.year, now.month)
    ws = wb["Tranzaksiyalar"]

    is_income = "Kirim" in data["type"]
    fill = PatternFill("solid", start_color="E2EFDA" if is_income else "FCE4D6")
    font_color = "375623" if is_income else "843C0C"

    currency = data.get("currency", "UZS")
    amount_display = f'{data["amount"]:,.2f} $' if currency == "USD" else data["amount"]
    row_data = [
        data.get("date", date.today().strftime("%d.%m.%Y")),
        now.strftime("%H:%M"),
        data["type"],
        data["category"].split(" ", 1)[-1],
        amount_display,
        currency,
        data["description"],
    ]
    next_row = ws.max_row + 1
    for col, val in enumerate(row_data, 1):
        cell = ws.cell(row=next_row, column=col, value=val)
        cell.border = BORDER
        cell.font = Font(name="Arial", size=10, color=font_color if col in [3, 4, 5] else "000000")
        cell.fill = fill if col == 3 else PatternFill()
        if col == 5:
            cell.number_format = "#,##0"
            cell.alignment = Alignment(horizontal="right")
        elif col in [1, 2, 3]:
            cell.alignment = Alignment(horizontal="center")

    wb.save(path)
    return str(path)


def get_monthly_summary(year: int, month: int) -> dict | None:
    path = get_excel_path(year, month)
    if not path.exists():
        return None
    wb = openpyxl.load_workbook(path)
    ws = wb["Tranzaksiyalar"]
    total_in, total_out = 0, 0
    cats_out, cats_in = {}, {}
    count = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        tur = str(row[2] or "")
        kat = str(row[3] or "Boshqa")
        try:
            amount = float(str(row[4] or 0).replace(",", ""))
        except ValueError:
            continue
        count += 1
        if "Kirim" in tur:
            total_in += amount
            cats_in[kat] = cats_in.get(kat, 0) + amount
        else:
            total_out += amount
            cats_out[kat] = cats_out.get(kat, 0) + amount
    return {"total_in": total_in, "total_out": total_out,
            "net": total_in - total_out, "cats_in": cats_in,
            "cats_out": cats_out, "count": count}


def get_today_records() -> list:
    now = datetime.now()
    path = get_excel_path(now.year, now.month)
    if not path.exists():
        return []
    today_str = date.today().strftime("%d.%m.%Y")
    wb = openpyxl.load_workbook(path)
    ws = wb["Tranzaksiyalar"]
    records = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] == today_str:
            records.append(row)
    return records


# ─── Telegram handlers ────────────────────────────────────────────────────────
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        ["➕ Xarajat qo'shish", "💰 Daromad qo'shish"],
        ["📊 Bu oy statistikasi", "📅 O'tgan oy hisoboti"],
        ["📋 Bugungi yozuvlar", "❓ Yordam"],
    ]
    await update.message.reply_text(
        "👋 *Salom! Moliya botiga xush kelibsiz!*\n\n"
        "Kunlik xarajat va daromadlaringizni yozing.\n"
        "Barchasi Excel fayliga avtomatik saqlanadi. 📊\n\n"
        "📁 Fayllar: `moliya_data/` papkasida",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True),
    )


async def handle_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    if text == "➕ Xarajat qo'shish":
        context.user_data["transaction_type"] = "🔴 Chiqim"
        return await _show_categories(update, CHIQIM_KATEGORIYALAR, "💸 Xarajat kategoriyasini tanlang:")
    elif text == "💰 Daromad qo'shish":
        context.user_data["transaction_type"] = "💚 Kirim"
        return await _show_categories(update, KIRIM_KATEGORIYALAR, "💰 Daromad kategoriyasini tanlang:")
    elif text == "📊 Bu oy statistikasi":
        await _current_stats(update)
    elif text == "📅 O'tgan oy hisoboti":
        await _last_month(update)
    elif text == "📋 Bugungi yozuvlar":
        await _today(update)
    elif text == "❓ Yordam":
        await _help(update)
    return ConversationHandler.END


async def _show_categories(update, categories, prompt):
    cols = 2
    keyboard = []
    for i in range(0, len(categories), cols):
        row = [InlineKeyboardButton(c, callback_data=f"cat:{c}") for c in categories[i:i+cols]]
        keyboard.append(row)
    keyboard.append([InlineKeyboardButton("❌ Bekor", callback_data="cancel")])
    await update.message.reply_text(prompt, reply_markup=InlineKeyboardMarkup(keyboard))
    return SELECTING_CATEGORY


async def category_selected(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "cancel":
        await query.edit_message_text("❌ Bekor qilindi.")
        return ConversationHandler.END
    context.user_data["category"] = query.data.replace("cat:", "")
    await query.edit_message_text(
        f"✅ Kategoriya: *{context.user_data['category']}*\n\n"
        f"💵 Summani kiriting (UZS):\n_Masalan: 50000_",
        parse_mode="Markdown"
    )
    return ENTERING_AMOUNT


async def amount_entered(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip().replace(",", "").replace(" ", "")
    # Valyuta turini aniqlash
    currency = "UZS"
    if "$" in text or "usd" in text.lower() or "dollar" in text.lower():
        currency = "USD"
    # Belgilarni olib tashlash
    text_clean = text.replace("$", "").replace("usd", "").replace("USD", "").replace("uzs", "").replace("UZS", "").strip()
    try:
        amount = float(text_clean)
        if amount <= 0:
            raise ValueError
    except ValueError:
        await update.message.reply_text(
            "❌ Faqat raqam kiriting:\n_Masalan: 50000 yoki 540$ yoki 100 USD_",
            parse_mode="Markdown"
        )
        return ENTERING_AMOUNT
    context.user_data["amount"] = amount
    context.user_data["currency"] = currency
    if currency == "USD":
        display = f"{amount:,.2f} $"
    else:
        display = f"{int(amount):,} UZS"
    await update.message.reply_text(
        f"✅ Summa: *{display}*\n\n📝 Izoh kiriting:\n_Masalan: Tushlik, Maosh, Dorixona_",
        parse_mode="Markdown"
    )
    return ENTERING_DESCRIPTION


async def description_entered(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["description"] = update.message.text.strip()
    context.user_data["date"] = date.today().strftime("%d.%m.%Y")
    keyboard = [
        [InlineKeyboardButton(f"📅 Bugun ({date.today().strftime('%d.%m')})", callback_data="date:today"),
         InlineKeyboardButton(f"📅 Kecha ({(date.today()-timedelta(1)).strftime('%d.%m')})", callback_data="date:yesterday")],
        [InlineKeyboardButton("✏️ Boshqa sana kiriting", callback_data="date:custom")],
    ]
    await update.message.reply_text("📅 Sana tanlang:", reply_markup=InlineKeyboardMarkup(keyboard))
    return SELECTING_DATE


async def date_selected(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "date:today":
        context.user_data["date"] = date.today().strftime("%d.%m.%Y")
    elif query.data == "date:yesterday":
        context.user_data["date"] = (date.today() - timedelta(1)).strftime("%d.%m.%Y")
    elif query.data == "date:custom":
        await query.edit_message_text("📅 Sanani kiriting (KK.OO.YYYY):\n_Masalan: 15.05.2026_",
                                       parse_mode="Markdown")
        context.user_data["waiting_custom_date"] = True
        return SELECTING_DATE
    return await _show_confirm_query(query, context)


async def custom_date_msg(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not context.user_data.get("waiting_custom_date"):
        return SELECTING_DATE
    try:
        d = datetime.strptime(update.message.text.strip(), "%d.%m.%Y")
        context.user_data["date"] = d.strftime("%d.%m.%Y")
        context.user_data["waiting_custom_date"] = False
    except ValueError:
        await update.message.reply_text("❌ Format noto'g'ri. KK.OO.YYYY:\n_Masalan: 15.05.2026_",
                                         parse_mode="Markdown")
        return SELECTING_DATE
    d = context.user_data
    keyboard = [[InlineKeyboardButton("✅ Saqlash", callback_data="confirm:yes"),
                 InlineKeyboardButton("❌ Bekor", callback_data="confirm:no")]]
    await update.message.reply_text(
        _confirm_text(d), parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )
    return CONFIRMING


def _confirm_text(d):
    currency = d.get("currency", "UZS")
    amount_str = f"{d['amount']:,.2f} $" if currency == "USD" else f"{int(d['amount']):,} UZS"
    return (f"📋 *Tasdiqlang:*\n\n"
            f"{'Tur:':<14} {d['transaction_type']}\n"
            f"{'Kategoriya:':<14} {d['category']}\n"
            f"{'Summa:':<14} *{amount_str}*\n"
            f"{'Izoh:':<14} {d['description']}\n"
            f"{'Sana:':<14} {d['date']}")


async def _show_confirm_query(query, context):
    d = context.user_data
    keyboard = [[InlineKeyboardButton("✅ Saqlash", callback_data="confirm:yes"),
                 InlineKeyboardButton("❌ Bekor", callback_data="confirm:no")]]
    await query.edit_message_text(_confirm_text(d), parse_mode="Markdown",
                                   reply_markup=InlineKeyboardMarkup(keyboard))
    return CONFIRMING


async def confirm_transaction(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "confirm:no":
        await query.edit_message_text("❌ Bekor qilindi.")
        return ConversationHandler.END
    d = context.user_data
    try:
        file_path = add_transaction({
            "type": d["transaction_type"],
            "category": d["category"],
            "amount": d["amount"],
            "currency": d.get("currency", "UZS"),
            "description": d["description"],
            "date": d["date"],
        })
        emoji = "💚" if "Kirim" in d["transaction_type"] else "🔴"
        file_name = Path(file_path).name
        currency = d.get("currency", "UZS")
        amount_str = f"{d['amount']:,.2f} $" if currency == "USD" else f"{int(d['amount']):,} UZS"
        await query.edit_message_text(
            f"✅ *Saqlandi!*\n\n"
            f"{emoji} *{amount_str}* — {d['description']}\n"
            f"📁 {d['category']}\n"
            f"📅 {d['date']}\n\n"
            f"💾 Fayl: `{file_name}`",
            parse_mode="Markdown"
        )
    except Exception as e:
        logger.error(f"Excel xatosi: {e}")
        await query.edit_message_text(f"❌ Xato: {str(e)}")
    context.user_data.clear()
    return ConversationHandler.END


async def _current_stats(update):
    now = datetime.now()
    msg = await update.message.reply_text("⏳ Hisoblanmoqda...")
    s = get_monthly_summary(now.year, now.month)
    if not s or s["count"] == 0:
        await msg.edit_text("📭 Bu oy hali yozuv yo'q.\n\nXarajat yoki daromad qo'shing!")
        return
    net_e = "✅" if s["net"] >= 0 else "⚠️"
    text = (f"📊 *{now.year}-yil {MONTH_NAMES[now.month]} oyi*\n\n"
            f"💚 Jami kirim:  *{s['total_in']:,.0f} UZS*\n"
            f"🔴 Jami chiqim: *{s['total_out']:,.0f} UZS*\n"
            f"{net_e} Sof balans:  *{s['net']:,.0f} UZS*\n"
            f"📝 Yozuvlar: {s['count']} ta\n")
    if s["cats_out"]:
        text += "\n🔴 *Eng katta chiqimlar:*\n"
        for k, v in sorted(s["cats_out"].items(), key=lambda x: -x[1])[:5]:
            text += f"  • {k}: {v:,.0f} UZS\n"
    await msg.edit_text(text, parse_mode="Markdown")


async def _last_month(update):
    now = datetime.now()
    year, month = (now.year - 1, 12) if now.month == 1 else (now.year, now.month - 1)
    msg = await update.message.reply_text("⏳ Yuklanmoqda...")
    s = get_monthly_summary(year, month)
    if not s or s["count"] == 0:
        await msg.edit_text(f"📭 {MONTH_NAMES[month]} oyi uchun ma'lumot yo'q.")
        return
    net_e = "✅" if s["net"] >= 0 else "⚠️"
    text = (f"📅 *{year}-yil {MONTH_NAMES[month]} oyi*\n\n"
            f"💚 Jami kirim:  *{s['total_in']:,.0f} UZS*\n"
            f"🔴 Jami chiqim: *{s['total_out']:,.0f} UZS*\n"
            f"{net_e} Sof balans:  *{s['net']:,.0f} UZS*\n"
            f"📝 Jami: {s['count']} ta\n")
    if s["cats_out"]:
        text += "\n🔴 *Chiqimlar:*\n"
        for k, v in sorted(s["cats_out"].items(), key=lambda x: -x[1])[:6]:
            pct = v / (s["total_out"] or 1) * 100
            text += f"  • {k}: {v:,.0f} ({pct:.0f}%)\n"
    await msg.edit_text(text, parse_mode="Markdown")


async def _today(update):
    records = get_today_records()
    today_str = date.today().strftime("%d.%m.%Y")
    if not records:
        await update.message.reply_text(f"📭 {today_str} — bugun hali yozuv yo'q.")
        return
    tin = sum(float(str(r[4] or 0).replace(",","")) for r in records if "Kirim" in str(r[2] or ""))
    tout = sum(float(str(r[4] or 0).replace(",","")) for r in records if "Chiqim" in str(r[2] or ""))
    text = f"📋 *{today_str} — Bugungi yozuvlar*\n\n"
    for r in records:
        e = "💚" if "Kirim" in str(r[2] or "") else "🔴"
        a = float(str(r[4] or 0).replace(",",""))
        text += f"{e} {r[3]} — *{a:,.0f} UZS* — _{r[5]}_\n"
    text += f"\n💚 Kirim: {tin:,.0f} | 🔴 Chiqim: {tout:,.0f} UZS"
    await update.message.reply_text(text, parse_mode="Markdown")


async def _help(update):
    await update.message.reply_text(
        "❓ *Yordam*\n\n"
        "➕ *Xarajat/Daromad qo'shish:*\n"
        "Kategoriya → Summa → Izoh → Sana → Tasdiqlash\n\n"
        "📊 *Ko'rsatkichlar:*\n"
        "• Bu oy statistikasi\n"
        "• O'tgan oy hisoboti\n"
        "• Bugungi yozuvlar\n\n"
        "💾 *Fayllar:* `moliya_data/` papkasida\n"
        "Har oy alohida Excel fayl yaratiladi.\n"
        "Masalan: `2026-05-moliya.xlsx`",
        parse_mode="Markdown"
    )


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("❌ Bekor qilindi.")
    return ConversationHandler.END


async def monthly_report_job(context: ContextTypes.DEFAULT_TYPE):
    now = datetime.now()
    year, month = (now.year - 1, 12) if now.month == 1 else (now.year, now.month - 1)
    s = get_monthly_summary(year, month)
    if not s or s["count"] == 0:
        return
    net_e = "✅" if s["net"] >= 0 else "⚠️"
    for uid in context.bot_data.get("users", set()):
        try:
            await context.bot.send_message(
                chat_id=uid,
                text=(f"🗓 *{year}-yil {MONTH_NAMES[month]} oyi tugadi!*\n\n"
                      f"💚 Kirim: *{s['total_in']:,.0f} UZS*\n"
                      f"🔴 Chiqim: *{s['total_out']:,.0f} UZS*\n"
                      f"{net_e} Balans: *{s['net']:,.0f} UZS*\n\n"
                      f"Yangi oy muborak! 💪"),
                parse_mode="Markdown"
            )
        except Exception as e:
            logger.error(f"Hisobot xatosi: {e}")


async def reset_month(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Joriy oy Excel faylini tozalash"""
    now = datetime.now()
    path = EXCEL_DIR / f"{now.year}-{now.month:02d}-moliya.xlsx"
    if path.exists():
        path.unlink()
        await update.message.reply_text(
            "🗑 *Joriy oy ma'lumotlari tozalandi!*\n\nEndi yangi yozuvlar qo'shishingiz mumkin.",
            parse_mode="Markdown"
        )
    else:
        await update.message.reply_text("📭 Bu oyda hali yozuv yo'q edi.")


async def track_user(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "users" not in context.bot_data:
        context.bot_data["users"] = set()
    context.bot_data["users"].add(update.effective_user.id)


def main():
    app = Application.builder().token(BOT_TOKEN).build()

    conv = ConversationHandler(
        entry_points=[MessageHandler(
            filters.Regex("^(➕ Xarajat qo'shish|💰 Daromad qo'shish)$"), handle_menu)],
        states={
            SELECTING_CATEGORY: [CallbackQueryHandler(category_selected)],
            ENTERING_AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, amount_entered)],
            ENTERING_DESCRIPTION: [MessageHandler(filters.TEXT & ~filters.COMMAND, description_entered)],
            SELECTING_DATE: [
                CallbackQueryHandler(date_selected, pattern="^date:"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, custom_date_msg),
            ],
            CONFIRMING: [CallbackQueryHandler(confirm_transaction, pattern="^confirm:")],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(MessageHandler(filters.ALL, track_user), group=-1)
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("reset", reset_month))
    app.add_handler(conv)
    app.add_handler(MessageHandler(
        filters.Regex("^(📊 Bu oy statistikasi|📅 O'tgan oy hisoboti|📋 Bugungi yozuvlar|❓ Yordam)$"),
        handle_menu
    ))

    # Har oy 1-si soat 09:00 da hisobot
    from datetime import time as dtime
    app.job_queue.run_monthly(
        monthly_report_job,
        when=dtime(9, 0, 0),
        day=1
    )

    logger.info("✅ Moliya boti ishga tushdi!")
    logger.info(f"📁 Excel fayllar: {EXCEL_DIR.absolute()}")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()

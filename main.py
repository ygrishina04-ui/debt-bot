import os
import tempfile
import pandas as pd
from flask import Flask, request
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

BOT_TOKEN = os.environ.get("BOT_TOKEN")
ADMIN_CHAT_ID = os.environ.get("ADMIN_CHAT_ID")

app = Flask(__name__)
telegram_app = Application.builder().token(BOT_TOKEN).build()


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! Я бот по дебиторской задолженности.\n\n"
        "Команда: /дебиторка"
    )


async def debt_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Загрузите Excel-файл с дебиторской задолженностью.\n\n"
        "Минимальные колонки:\n"
        "Клиент | Сумма задолженности"
    )


async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document

    if not document:
        return

    file_name = document.file_name or ""

    if not file_name.endswith((".xlsx", ".xls")):
        await update.message.reply_text("Пожалуйста, загрузите файл Excel: .xlsx или .xls")
        return

    await update.message.reply_text("Файл получила. Начинаю проверку...")

    tg_file = await context.bot.get_file(document.file_id)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        file_path = tmp.name

    await tg_file.download_to_drive(file_path)

    try:
        df = pd.read_excel(file_path)

        required_columns = ["Клиент", "Сумма задолженности"]
        missing = [col for col in required_columns if col not in df.columns]

        if missing:
            await update.message.reply_text(
                "В файле не хватает колонок:\n" + "\n".join(missing)
            )
            return

        df = df[df["Сумма задолженности"].fillna(0) > 0]

        total_clients = len(df)
        total_sum = df["Сумма задолженности"].sum()

        message = (
            "Файл прочитан ✅\n\n"
            f"Клиентов с задолженностью: {total_clients}\n"
            f"Общая сумма: {total_sum:,.2f} руб.\n\n"
            "Следующий шаг: подключим базу клиентов и подтягивание email."
        )

        await update.message.reply_text(message.replace(",", " "))

    except Exception as e:
        await update.message.reply_text(f"Ошибка при чтении файла:\n{e}")

    finally:
        try:
            os.remove(file_path)
        except Exception:
            pass


telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(CommandHandler("дебиторка", debt_start))
telegram_app.add_handler(MessageHandler(filters.Document.ALL, handle_file))


@app.route("/")
def index():
    return "Debt bot is running"


@app.route("/webhook", methods=["POST"])
async def webhook():
    update = Update.de_json(request.get_json(force=True), telegram_app.bot)
    await telegram_app.process_update(update)
    return "ok"


if __name__ == "__main__":
    import asyncio

    async def main():
        await telegram_app.initialize()
        await telegram_app.start()
        app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))

    asyncio.run(main())

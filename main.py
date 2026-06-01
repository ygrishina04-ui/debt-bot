import os
import tempfile
from flask import Flask, request
import requests
from openpyxl import load_workbook

app = Flask(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN")
ADMIN_CHAT_ID = os.environ.get("ADMIN_CHAT_ID")


@app.route("/")
def index():
    return "Debt bot is running"


def send_message(chat_id, text):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": chat_id,
        "text": text
    }
    requests.post(url, json=payload)


def get_telegram_file(file_id):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/getFile"
    response = requests.get(url, params={"file_id": file_id})
    data = response.json()

    if not data.get("ok"):
        raise Exception("Не удалось получить файл из Telegram")

    file_path = data["result"]["file_path"]
    file_url = f"https://api.telegram.org/file/bot{BOT_TOKEN}/{file_path}"

    file_response = requests.get(file_url)
    file_response.raise_for_status()

    return file_response.content


def analyze_excel(file_content):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(file_content)
        tmp_path = tmp.name

    try:
        wb = load_workbook(tmp_path, data_only=True)
        ws = wb.active

        headers = [cell.value for cell in ws[1]]

        if "Клиент" not in headers or "Сумма задолженности" not in headers:
            return (
                "В файле должны быть колонки:\n"
                "Клиент\n"
                "Сумма задолженности"
            )

        client_idx = headers.index("Клиент") + 1
        sum_idx = headers.index("Сумма задолженности") + 1

        total_clients = 0
        total_sum = 0
        preview = []

        for row in range(2, ws.max_row + 1):
            client = ws.cell(row=row, column=client_idx).value
            amount = ws.cell(row=row, column=sum_idx).value

            if not client or amount in [None, ""]:
                continue

            try:
                amount_num = float(str(amount).replace(" ", "").replace(",", "."))
            except Exception:
                continue

            if amount_num <= 0:
                continue

            total_clients += 1
            total_sum += amount_num

            if len(preview) < 10:
                preview.append(f"• {client}: {amount_num:,.2f} руб.".replace(",", " "))

        if total_clients == 0:
            return "Файл прочитан, но задолженность не найдена."

        message = (
            "Файл прочитан ✅\n\n"
            f"Клиентов с задолженностью: {total_clients}\n"
            f"Общая сумма: {total_sum:,.2f} руб.\n\n"
            "Первые строки:\n"
            + "\n".join(preview)
            + "\n\nСледующий шаг: подключим базу клиентов и email."
        )

        return message.replace(",", " ")

    finally:
        try:
            os.remove(tmp_path)
        except Exception:
            pass


@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.get_json()

    if not data:
        return "ok"

    message = data.get("message", {})
    chat = message.get("chat", {})
    chat_id = chat.get("id")
    text = message.get("text", "")
    document = message.get("document")

    if not chat_id:
        return "ok"

    if text == "/start":
        send_message(chat_id, "Привет! Я бот по дебиторской задолженности.\n\nКоманда: /дебиторка")
        return "ok"

    if text == "/дебиторка":
        send_message(
            chat_id,
            "Загрузите Excel-файл с дебиторской задолженностью.\n\n"
            "Обязательные колонки:\n"
            "Клиент | Сумма задолженности"
        )
        return "ok"

    if document:
        file_name = document.get("file_name", "")

        if not file_name.endswith((".xlsx", ".xls")):
            send_message(chat_id, "Пожалуйста, загрузите Excel-файл .xlsx или .xls")
            return "ok"

        try:
            send_message(chat_id, "Файл получила. Начинаю проверку...")
            file_content = get_telegram_file(document["file_id"])
            result = analyze_excel(file_content)
            send_message(chat_id, result)

        except Exception as e:
            send_message(chat_id, f"Ошибка при обработке файла:\n{e}")

        return "ok"

    send_message(chat_id, "Я получил сообщение ✅\n\nПока понимаю команду: /дебиторка")
    return "ok"

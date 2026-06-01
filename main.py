import os
import tempfile
from flask import Flask, request
import requests
from openpyxl import load_workbook
import gspread
from google.oauth2.service_account import Credentials

app = Flask(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN")
ADMIN_CHAT_ID = os.environ.get("ADMIN_CHAT_ID")

GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID")
GOOGLE_SERVICE_ACCOUNT_EMAIL = os.environ.get("GOOGLE_SERVICE_ACCOUNT_EMAIL")
GOOGLE_PRIVATE_KEY = os.environ.get("GOOGLE_PRIVATE_KEY")


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


def get_google_sheet():
    private_key = GOOGLE_PRIVATE_KEY.replace("\\n", "\n")

    info = {
        "type": "service_account",
        "client_email": GOOGLE_SERVICE_ACCOUNT_EMAIL,
        "private_key": private_key,
        "token_uri": "https://oauth2.googleapis.com/token"
    }

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = Credentials.from_service_account_info(info, scopes=scopes)
    client = gspread.authorize(creds)

    return client.open_by_key(GOOGLE_SHEET_ID)


def get_clients_base():
    sheet = get_google_sheet()
    ws = sheet.worksheet("БАЗА_КЛИЕНТОВ")

    rows = ws.get_all_records()

    base = {}

    for row in rows:
        client = str(row.get("Клиент", "")).strip()
        email = str(row.get("Почта", "")).strip()

        if client:
            base[normalize_client(client)] = {
                "client": client,
                "email": email
            }

    return base


def normalize_client(name):
    return (
        str(name)
        .strip()
        .lower()
        .replace("«", "")
        .replace("»", "")
        .replace('"', "")
        .replace("ооо ", "")
        .replace("оoo ", "")
        .replace("ип ", "")
        .replace("  ", " ")
    )


def analyze_excel(file_content):
    clients_base = get_clients_base()

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

        ready = []
        no_email = []
        not_found = []

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

            norm_client = normalize_client(client)
            base_item = clients_base.get(norm_client)

            if not base_item:
                not_found.append((client, amount_num))
                continue

            email = base_item.get("email", "")

            if not email:
                no_email.append((client, amount_num))
                continue

            ready.append((client, amount_num, email))

        result = (
            "Файл прочитан ✅\n\n"
            f"Клиентов с задолженностью: {total_clients}\n"
            f"Общая сумма: {total_sum:,.2f} руб.\n\n"
            f"✅ Готово к рассылке: {len(ready)}\n"
            f"⚠️ Клиент найден, но нет почты: {len(no_email)}\n"
            f"❌ Клиент не найден в базе: {len(not_found)}\n"
        ).replace(",", " ")

        if ready:
            result += "\n\nПервые готовые к рассылке:\n"
            for client, amount, email in ready[:10]:
                result += f"• {client}: {amount:,.2f} руб. → {email}\n".replace(",", " ")

        if no_email:
            result += "\n\n⚠️ Нет почты:\n"
            for client, amount in no_email[:10]:
                result += f"• {client}: {amount:,.2f} руб.\n".replace(",", " ")

        if not_found:
            result += "\n\n❌ Нет в базе:\n"
            for client, amount in not_found[:10]:
                result += f"• {client}: {amount:,.2f} руб.\n".replace(",", " ")

        result += "\n\nСледующий шаг: добавим предпросмотр писем и кнопку подтверждения."

        return result

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
            "Клиент | Сумма задолженности\n\n"
            "Почта подтянется из листа БАЗА_КЛИЕНТОВ."
        )
        return "ok"

    if document:
        file_name = document.get("file_name", "")

        if not file_name.endswith((".xlsx", ".xls")):
            send_message(chat_id, "Пожалуйста, загрузите Excel-файл .xlsx или .xls")
            return "ok"

        try:
            send_message(chat_id, "Файл получила. Сверяю с базой клиентов...")
            file_content = get_telegram_file(document["file_id"])
            result = analyze_excel(file_content)
            send_message(chat_id, result)

        except Exception as e:
            send_message(chat_id, f"Ошибка при обработке файла:\n{e}")

        return "ok"

    send_message(chat_id, "Я получил сообщение ✅\n\nПока понимаю команду: /дебиторка")
    return "ok"

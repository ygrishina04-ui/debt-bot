import os
from flask import Flask, request
import requests

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


@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.get_json()

    if not data:
        return "ok"

    message = data.get("message", {})
    chat = message.get("chat", {})
    chat_id = chat.get("id")
    text = message.get("text", "")

    if not chat_id:
        return "ok"

    if text == "/start":
        send_message(chat_id, "Привет! Я бот по дебиторской задолженности.\n\nКоманда: /дебиторка")
        return "ok"

    if text == "/дебиторка":
        send_message(chat_id, "Загрузите Excel-файл с дебиторской задолженностью.")
        return "ok"

    send_message(chat_id, "Я получил сообщение ✅\n\nПока понимаю команду: /дебиторка")
    return "ok"

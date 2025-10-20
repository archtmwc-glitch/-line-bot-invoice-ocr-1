from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, ImageMessage, TextSendMessage
import pytesseract
from PIL import Image
import openpyxl
import os

app = Flask(__name__)

# LINE Bot credentials
CHANNEL_ACCESS_TOKEN = 'YOUR_CHANNEL_ACCESS_TOKEN'
CHANNEL_SECRET = 'YOUR_CHANNEL_SECRET'

line_bot_api = LineBotApi(CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(CHANNEL_SECRET)

# Excel file path
EXCEL_FILE = 'invoice_data.xlsx'

# Ensure Excel file exists
if not os.path.exists(EXCEL_FILE):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "發票資料"
    ws.append(["發票號碼", "金額", "圖片名稱"])
    wb.save(EXCEL_FILE)

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return 'OK'

@handler.add(MessageEvent, message=ImageMessage)
def handle_image_message(event):
    message_id = event.message.id
    image_content = line_bot_api.get_message_content(message_id)

    image_path = f"{message_id}.jpg"
    with open(image_path, "wb") as f:
        for chunk in image_content.iter_content():
            f.write(chunk)

    # OCR processing
    image = Image.open(image_path)
    text = pytesseract.image_to_string(image, lang='eng+chi_tra')

    # Extract invoice number and amount
    invoice_number = ""
    amount = ""
    for line in text.splitlines():
        if "SA-" in line or "發票號碼" in line:
            invoice_number = line.strip()
        if "總計" in line or "金額" in line:
            amount = ''.join(filter(str.isdigit, line))

    # Update Excel
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    ws.append([invoice_number, amount, image_path])
    wb.save(EXCEL_FILE)

    # Reply to user
    reply = f"發票號碼：{invoice_number}"
金額：{amount} 元"
    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))

if __name__ == "__main__":
    app.run()

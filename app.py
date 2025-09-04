import os, json, re
from datetime import datetime, timezone, timedelta
from flask import Flask, request, abort, jsonify
from linebot.v3 import WebhookParser
from linebot.v3.messaging import MessagingApi, Configuration, ApiClient
from linebot.v3.webhooks import MessageEvent, TextMessageContent
import gspread
from google.oauth2.service_account import Credentials

# ========= 環境變數 =========
CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET", "")
CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN", "")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "")
SHEET_NAME = os.getenv("SHEET_NAME", "records")
TIMEZONE_HOURS = int(os.getenv("TIMEZONE_HOURS", "8"))
SA_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
INTERNAL_SHIP_SECRET = os.getenv("INTERNAL_SHIP_SECRET", "").strip()  # Apps Script 內部密鑰

if not CHANNEL_SECRET or not CHANNEL_ACCESS_TOKEN or not SPREADSHEET_ID or not SA_JSON:
    raise RuntimeError("請在雲端後台設定：LINE_CHANNEL_SECRET、LINE_CHANNEL_ACCESS_TOKEN、SPREADSHEET_ID、GOOGLE_SERVICE_ACCOUNT_JSON")

# ========= Google Sheets 連線 =========
SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]

def gs_client():
    raw = SA_JSON
    if not raw:
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON is empty")
    try:
        if raw.startswith("{"):
            info = json.loads(raw)
        else:
            import base64
            info = json.loads(base64.b64decode(raw).decode("utf-8"))
    except Exception as e:
        raise RuntimeError(f"Invalid GOOGLE_SERVICE_ACCOUNT_JSON: {e}")
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)

def ensure_sheet(gc, spreadsheet_id, sheet_name):
    sh = gc.open_by_key(spreadsheet_id)
    try:
        ws = sh.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=sheet_name, rows=1000, cols=20)
        ws.update('A1:K1', [[
            "Timestamp","UserId","DisplayName",
            "Item","Quantity","Price","Amount","Note","RawText",
            "通知出貨","通知狀態"
        ]])
    return ws

# ========= 文字解析 =========
ALIASES = {
  "item":  ["品項","項目","品名","item"],
  "qty":   ["數量","數目","qty","數"],
  "price": ["單價","價格","價","price"],
  "note":  ["備註","備註說明","note"],
}
def norm(k):
    k=k.strip().lower()
    for std,arr in ALIASES.items():
        for a in arr:
            if k==a.lower(): return std
    return None

def parse(text):
    if not re.match(r"^\s*[紀记]錄[:：\s]", text): return None
    body = re.split(r"[：:]", text, 1)[1].strip()
    parts = re.split(r"[,，;；\n]+", body)
    r = {"item":None,"qty":None,"price":None,"note":None}
    for p in parts:
        if not p.strip(): continue
        kv = re.split(r"[=：:]", p, 1)
        if len(kv)==2:
            k,v = kv[0].strip(), kv[1].strip()
            std = norm(k)
            if std: r[std]=v
        else:
            r["note"] = (r["note"]+" | " if r["note"] else "") + p.strip()
    def num(s):
        if s is None: return None
        s=s.replace(",","").strip()
        try: return float(s) if "." in s else int(s)
        except: return None
    r["qty"]=num(r["qty"]); r["price"]=num(r["price"])
    return r

def now_str():
    return datetime.now(timezone(timedelta(hours=TIMEZONE_HOURS))).strftime("%Y-%m-%d %H:%M:%S")

# ========= Flask + LINE =========
app = Flask(__name__)
parser = WebhookParser(CHANNEL_SECRET)
cfg = Configuration(access_token=CHANNEL_ACCESS_TOKEN)
api_client = ApiClient(cfg)
msg_api = MessagingApi(api_client)

@app.get("/healthz")
def healthz():
    return "ok", 200

# 自我測試：直接寫一筆資料到表格
@app.get("/debug/write")
def debug_write():
    try:
        gc = gs_client()
        ws = ensure_sheet(gc, SPREADSHEET_ID, SHEET_NAME)
        ws.append_row(
            [now_str(), "debug", "", "測試筆", 1, 2, 2, "from /debug/write", "—", "", ""],
            value_input_option="USER_ENTERED"
        )
        return "ok: wrote a test row", 200
    except Exception as e:
        import traceback, io
        buf = io.StringIO()
        traceback.print_exc(file=buf)
        return f"error: {e}\n\n{buf.getvalue()}", 500

# LINE Webhook：解析訊息 → 寫入 Sheets → 回覆
@app.post("/line/webhook")
def webhook():
    sig = request.headers.get("X-Line-Signature","")
    body = request.get_data(as_text=True)
    try:
        events = parser.parse(body, sig)
    except Exception:
        abort(400)

    # 準備 Sheets
    try:
        gc = gs_client()
        ws = ensure_sheet(gc, SPREADSHEET_ID, SHEET_NAME)
    except Exception as e:
        print("Sheets 初始化失敗:", e)
        return "OK"

    for ev in events:
        if isinstance(ev, MessageEvent) and isinstance(ev.message, TextMessageContent):
            t = ev.message.text.strip()
            data = parse(t)
            if not data:
                try:
                    msg_api.reply_message(
                        reply_token=ev.reply_token,
                        messages=[{"type":"text","text":"要記錄到雲端表格，請用：\n紀錄：品項=蘋果, 數量=10, 單價=50, 備註=特價"}]
                    )
                except Exception as e:
                    print("reply_message failed:", e)
                continue

            qty, price = data.get("qty"), data.get("price")
            amount = qty*price if isinstance(qty,(int,float)) and isinstance(price,(int,float)) else ""
            user_id = getattr(ev.source, "user_id", None)
            try:
                ws.append_row([
                    now_str(), user_id or "", "",
                    data.get("item") or "", qty or "", price or "", amount,
                    data.get("note") or "", t, "", ""
                ], value_input_option="USER_ENTERED")
                # 回覆成功訊息
                try:
                    msg_api.reply_message(
                        reply_token=ev.reply_token,
                        messages=[{"type":"text","text":f"✅ 已記錄\n品項：{data.get('item') or '-'}\n數量：{data.get('qty') or '-'}\n單價：{data.get('price') or '-'}\n備註：{data.get('note') or '-'}"}]
                    )
                except Exception as e:
                    print("reply_message failed:", e)
            except Exception as e:
                print("寫入/回覆失敗:", e)
    return "OK"

# 出貨通知 API（給 Apps Script 呼叫）
@app.post("/ship/notify")
def ship_notify():
    # 驗證密鑰（簡單保護）
    if INTERNAL_SHIP_SECRET:
        chk = request.headers.get("X-Internal-Secret", "")
        if chk != INTERNAL_SHIP_SECRET:
            return "forbidden", 403

    data = request.get_json(silent=True) or {}
    sheet_id = data.get("spreadsheet_id") or SPREADSHEET_ID
    sheet_name = data.get("sheet_name") or SHEET_NAME
    row = int(data.get("row") or 0)
    custom_msg = data.get("message")

    if row <= 1:
        return "bad row", 400

    try:
        gc = gs_client()
        sh = gc.open_by_key(sheet_id)
        ws = sh.worksheet(sheet_name)

        values = ws.row_values(row)
        def _get(col_idx):
            return values[col_idx-1] if len(values) >= col_idx else ""

        user_id = _get(2)   # B: UserId
        item    = _get(4)   # D: Item
        qty     = _get(5)   # E: Quantity
        price   = _get(6)   # F: Price
        flag    = _get(10)  # J: 通知出貨

        if flag != "出貨":
            return jsonify({"skip": True, "reason": "J欄不是『出貨』"}), 200
        if not user_id:
            return jsonify({"error": "該列沒有 UserId，無法推播"}), 400

        text = custom_msg or f"📦 您的訂單已出貨！\n品項：{item}\n數量：{qty}\n單價：{price}\n感謝您的購買。"
        try:
            msg_api.push_message(
                to=user_id,
                messages=[{"type": "text", "text": text}]
            )
        except Exception as e:
            print("push_message failed:", e)
            return jsonify({"error": f"push failed: {e}"}), 500

        # 回寫 K 欄：通知狀態
        ts = now_str()
        ws.update_cell(row, 11, f"已通知 {ts}")  # K = 11
        return jsonify({"ok": True, "row": row}), 200

    except Exception as e:
        import traceback, io
        buf = io.StringIO()
        traceback.print_exc(file=buf)
        return f"error: {e}\n\n{buf.getvalue()}", 500


if __name__ == "__main__":
    port = int(os.getenv("PORT", "8080"))
    app.run(host="0.0.0.0", port=port)

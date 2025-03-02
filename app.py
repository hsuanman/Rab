from flask import Flask, request, abort
from linebot.v3 import WebhookHandler
from linebot.v3.exceptions import InvalidSignatureError
from linebot.v3.webhooks import MessageEvent
from linebot.v3.messaging import Configuration, ApiClient, MessagingApi, ReplyMessageRequest, TextMessage
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os
import datetime

app = Flask(__name__)

# LINE Bot 設定
configuration = Configuration(access_token=os.getenv('CHANNEL_ACCESS_TOKEN'))
handler = WebhookHandler(os.getenv('CHANNEL_SECRET'))

# Google OAuth2 設定
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
REDIRECT_URI = 'https://4b55-36-231-192-94.ngrok-free.app/callback'  # 替換為你的實際 redirect_uri

# 初始化 OAuth2 Flow
flow = Flow.from_client_secrets_file(
    'client_secret.json',  # 從 Google Cloud Console 下載的憑證檔案
    scopes=SCOPES,
    redirect_uri=REDIRECT_URI
)

# 處理 LINE 的 Webhook
@app.route("/webhook", methods=['POST'])
def webhook():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return 'OK'

# 處理訊息事件
@handler.add(MessageEvent)
def handle_message(event):
    # 定義允許的關鍵字
    allowed_keywords = ['今日', 'W', 'O', '全部']

    # 檢查用戶輸入是否為允許的關鍵字
    if event.message.text in allowed_keywords:
        if event.message.text == '今日':
            events = get_google_calendar_events('今日')
        elif event.message.text == 'W':
            events = get_google_calendar_events('W')
        elif event.message.text == 'O':
            events = get_google_calendar_events('O')
        else:
            events = get_google_calendar_events('全部')

        if events:
            message = "\n".join([f"{i+1}. {format_datetime(event['start'].get('dateTime', event['start'].get('date')))} {event['summary']} " 
                               for i, event in enumerate(events)])
            with ApiClient(configuration) as api_client:
                line_bot_api = MessagingApi(api_client)
                line_bot_api.reply_message_with_http_info(
                    ReplyMessageRequest(
                        reply_token=event.reply_token,
                        messages=[TextMessage(text=message)]
                    )
                )
        else:
            with ApiClient(configuration) as api_client:
                line_bot_api = MessagingApi(api_client)
                line_bot_api.reply_message_with_http_info(
                    ReplyMessageRequest(
                        reply_token=event.reply_token,
                        messages=[TextMessage(text='沒有找到任何事件。')]
                    )
                )
    else:
        # 如果輸入的不是關鍵字，回傳提示訊息
        with ApiClient(configuration) as api_client:
            line_bot_api = MessagingApi(api_client)
            line_bot_api.reply_message_with_http_info(
                ReplyMessageRequest(
                    reply_token=event.reply_token,
                    messages=[TextMessage(text='請輸入關鍵字查詢（今日、W、O、全部）。')]
                )
            )

def is_today(event_time):
    # 將事件時間轉換為日期
    if 'T' in event_time:  # 如果是 dateTime 格式
        event_date = datetime.datetime.fromisoformat(event_time).date()
    else:  # 如果是 date 格式
        event_date = datetime.datetime.fromisoformat(event_time + 'T00:00:00+00:00').date()

    # 取得今天的日期（本地時間）
    today = datetime.datetime.now().date()

    # 判斷事件日期是否等於今天
    return event_date == today


# 取得 Google 行事曆事件
def get_google_calendar_events(filter_type):
    if not os.path.exists('token.json'):
        print("未找到 token.json 檔案")
        return None

    try:
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    except ValueError as e:
        print(f"載入權杖失敗：{e}")
        return None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("刷新權杖...")
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"刷新權杖失敗：{e}")
                return None
        else:
            print("權杖無效或過期")
            return None

    service = build('calendar', 'v3', credentials=creds)
    now = datetime.datetime.utcnow()
    time_min = now.isoformat() + 'Z'  # 從現在開始
    time_max = None
    #time_max = (now + datetime.timedelta(days=1)).isoformat() + 'Z'  # 只查詢未來一天的事件

    print(f"查詢日曆 ID: fce0ef2f6c8f301530dc62bd077eab1fd3bb14dd39d580e236bfa390362d6d29@group.calendar.google.com")
    print(f"時間範圍: {time_min} 到 {time_max}")

    try:
        events_result = service.events().list(
            calendarId='fce0ef2f6c8f301530dc62bd077eab1fd3bb14dd39d580e236bfa390362d6d29@group.calendar.google.com',
            timeMin=time_min,
            timeMax=time_max,
            maxResults=100,  # 最多 100 個事件
            singleEvents=True,
            orderBy='startTime'
        ).execute()

        events = events_result.get('items', [])
        print(f"找到 {len(events)} 個事件")

        # 根據 filter_type 過濾事件
        if filter_type == '今日':
            print("過濾今日的事件...")
            events = [event for event in events if is_today(event['start'].get('dateTime', event['start'].get('date')))]
            print(f"過濾後的事件數量: {len(events)}")
        elif filter_type == 'W':
            events = [event for event in events if 'W' in event['summary']]
        elif filter_type == 'O':
            events = [event for event in events if 'O' in event['summary']]

        return events
    except Exception as e:
        print(f"查詢行事曆失敗：{e}")
        return None


# 格式化時間為 MM DD 24hr
def format_datetime(event_time):
    dt = datetime.datetime.fromisoformat(event_time)
    return dt.strftime('%m-%d %H:%M')

# 處理 Google OAuth2 回調
@app.route("/callback")
def callback():
    code = request.args.get('code')
    if not code:
        return '缺少授權碼', 400

    try:
        # 使用授權碼取得權杖
        flow.fetch_token(code=code)
        creds = flow.credentials

        # 將權杖儲存起來（例如存到檔案中）
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

        return '授權成功！你可以關閉此頁面。'
    except Exception as e:
        print(f'取得權杖失敗：{e}')
        return '授權失敗', 500

if __name__ == "__main__":
    app.run(port=3000)
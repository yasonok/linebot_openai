from flask import Flask, request
import json
import pandas as pd
from linebot import LineBotApi, WebhookHandler
from linebot.models import MessageEvent, TextMessage, TextSendMessage, TemplateSendMessage, ButtonsTemplate, PostbackAction

app = Flask(__name__)
line_bot_api = LineBotApi('ZNlC5wzZ6N/l7ZH3N4QQi5YQNExXR0sQ9qTRK2dnfLGeRP+uQeq/mtCBSkByCjX8TsEaEXwY7mzQcP3Nv/bndlV4ApAhgoj8KzxJbI+E00zcg4tyEg9qfe8SlC5d657yvJ84DaEaacM84crmldRvzAdB04t89/1O/w1cDnyilFU=')
handler = WebhookHandler('231d25a1227792179a03cdf2641216e0')

options = {
    '機房位置': r'C:\Users\yason\OneDrive\桌面\py\03.xlsx',
    '通訊錄': r'C:\Users\yason\OneDrive\桌面\py\02.xlsx',
    '基站位置': r'C:\Users\yason\OneDrive\桌面\py\04.xlsx'
}

user_states = {}

def generate_buttons_template():
    buttons_template = ButtonsTemplate(
        title='請選擇選項',
        text='請選擇以下其中一個選項：',
        actions=[
            PostbackAction(label='機房位置', data='option_機房位置'),
            PostbackAction(label='通訊錄', data='option_通訊錄'),
            PostbackAction(label='基站位置', data='option_基站位置')
        ]
    )
    return TemplateSendMessage(alt_text='請選擇選項', template=buttons_template)

@app.route("/", methods=['POST'])
def linebot():
    body = request.get_data(as_text=True)
    json_data = json.loads(body)
    tk = json_data['events'][0]['replyToken']
    event_type = json_data['events'][0]['type']
    user_id = json_data['events'][0]['source']['userId']
    msg = json_data['events'][0]['message']['text'].strip().lower() if 'message' in json_data['events'][0] else None

    if event_type == 'postback':
        postback_data = json_data['events'][0]['postback']['data']
        if postback_data.startswith('option_'):
            selected_option = postback_data.replace('option_', '')
            if selected_option == '機房位置':
                excel_file_path = options[selected_option]
                df = pd.read_excel(excel_file_path, header=None)
                df[0] = df[0].str.upper()
                room_names = df[0].tolist()
                room_names_str = '\n'.join(room_names)
                line_bot_api.reply_message(tk, TextSendMessage(f"機房名稱列表：\n{room_names_str}\n請輸入您要查詢的機房名稱。"))
                user_states[user_id] = f'waiting_name_{selected_option}'
            else:
                line_bot_api.reply_message(tk, TextSendMessage(f"請輸入{selected_option}的名稱。"))
                user_states[user_id] = f'waiting_name_{selected_option}'
        return 'OK'

    elif event_type == 'message':
        if msg == 'reset':
            # 使用者輸入了 "reset"，重新跳出按鈕訊息
            line_bot_api.reply_message(tk, generate_buttons_template())
            user_states[user_id] = 'waiting_option'
        elif user_id not in user_states:
            line_bot_api.reply_message(tk, generate_buttons_template())
            user_states[user_id] = 'waiting_option'
        elif user_states[user_id] == 'waiting_option':
            selected_option = msg
            if selected_option in options:
                excel_file_path = options[selected_option]
                df = pd.read_excel(excel_file_path, header=None)
                df[0] = df[0].str.upper()
                room_names = df[0].tolist()
                room_names_str = '\n'.join(room_names)
                line_bot_api.reply_message(tk, TextSendMessage(f"機房名稱列表：\n{room_names_str}\n請輸入您要查詢的機房名稱。"))
                user_states[user_id] = f'waiting_name_{selected_option}'
            else:
                line_bot_api.reply_message(tk, TextSendMessage("請選擇有效的選項：機房位置、通訊錄、基站位置。"))
        elif user_states[user_id].startswith('waiting_name_'):
            selected_option = user_states[user_id].replace('waiting_name_', '')
            excel_file_path = options[selected_option]
            msg = msg.upper()
            try:
                df = pd.read_excel(excel_file_path, header=None)
                df[0] = df[0].str.upper()
                row = df.loc[df[0] == msg].iloc[0]
                if selected_option == '機房位置':
                    response_message = f"機房名稱：{row[0]}\n位置：{row[2]}"
                elif selected_option == '通訊錄':
                    response_message = f"工程師姓名：{row[0]}\n電話：{row[2]}\nMAIL：{row[4]}"
                elif selected_option == '基站位置':
                    response_message = f"基站名稱：{row[0]}\n基站名稱：{row[1]}\n地址：{row[2]}"
                else:
                    response_message = "找不到相應的資料。"
                line_bot_api.reply_message(tk, TextSendMessage(response_message))
                del user_states[user_id]
            except IndexError:
                line_bot_api.reply_message(tk, TextSendMessage(f"找不到名稱為{msg}的資訊。請重新輸入。"))
                user_states[user_id] = f'waiting_name_{selected_option}'

    return 'OK'

if __name__ == "__main__":
    app.run()

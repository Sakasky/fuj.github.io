from django.shortcuts import render
from django.http import HttpResponse, HttpResponseBadRequest, HttpResponseForbidden
from django.views.decorators.csrf import csrf_exempt
from django.conf import settings
from linebot import LineBotApi, WebhookParser
from linebot.exceptions import InvalidSignatureError, LineBotApiError
from linebot.models import MessageEvent, TextSendMessage
from liffpy import (LineFrontendFramework as LIFF,ErrorResponse)
from linebot.models.send_messages import QuickReply, QuickReplyButton, VideoSendMessage
from linebot.models.actions import Action, MessageAction, PostbackAction, URIAction
from linebot.models.events import Postback, PostbackEvent
from oauth2client.service_account import ServiceAccountCredentials
from typing import Counter
from django.shortcuts import render
from django.http import HttpResponse, HttpResponseBadRequest, HttpResponseForbidden
from django.views.decorators.csrf import csrf_exempt
from django.conf import settings
from linebot import LineBotApi, WebhookParser
from linebot.exceptions import InvalidSignatureError, LineBotApiError
from linebot.models import MessageEvent, TextSendMessage
from linebot.models import actions
from linebot.models.actions import Action, MessageAction, PostbackAction, URIAction, DatetimePickerAction
from linebot.models.events import Postback, PostbackEvent
from linebot.models import ImageCarouselTemplate, ImageCarouselColumn, URIAction, CarouselTemplate, CarouselColumn
from linebot.models.messages import ImageMessage, TextMessage, VideoMessage
from linebot.models.send_messages import QuickReply, QuickReplyButton, VideoSendMessage
from gspread.cell import Cell
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from linebot.models.send_messages import ImageSendMessage
import gspread
import os
import rfc3339
from linebot.models import (MessageEvent,TextSendMessage,TemplateSendMessage,ButtonsTemplate,
                            MessageTemplateAction,PostbackEvent,PostbackTemplateAction,ConfirmTemplate)
import datetime
import pandas as pd
import numpy as np

line_bot_api = LineBotApi(settings.LINE_CHANNEL_ACCESS_TOKEN)
parser = WebhookParser(settings.LINE_CHANNEL_SECRET)
global PATH 
# 請改成 sheet_cred的絕對路徑
PATH =  '/Users/jzhuang/Desktop/專案&接案/診所機器人/infolinebot/sheet_cred.json'
# 請改成 credentials.json的絕對路徑
global PATH2
PATH2 = '/Users/jzhuang/Desktop/專案&接案/診所機器人/infolinebot/credentials.json'


def getsheet_member(): # 患者基本資料表讀取
    scopes = ["https://spreadsheets.google.com/feeds"]
    global PATH
    cred_path = PATH
    credentials = ServiceAccountCredentials.from_json_keyfile_name(cred_path, scopes)
    client = gspread.authorize(credentials)
    sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1HFCXfq3XDNokyQI3F3_yi0Jpr8ie9TK16j-zOQTAAT4/edit#gid=1711565043").get_worksheet(0)
    return sheet

def getsheet_appoint(): # 預約門診單讀取
    scopes = ["https://spreadsheets.google.com/feeds"]
    global PATH
    cred_path = PATH
    credentials = ServiceAccountCredentials.from_json_keyfile_name(cred_path, scopes)
    client = gspread.authorize(credentials)
    sheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1xFUgSiqU08UqjkYeB-wv8GiVnYFeT1s1MfPySVqYhQg/edit#gid=1112453375').get_worksheet(0)
    return sheet

# def getsheet_lineid1(): # 身分證對應 
#     scopes = ["https://spreadsheets.google.com/feeds"]
#     global PATH
#     cred_path = PATH
#     credentials = ServiceAccountCredentials.from_json_keyfile_name(cred_path, scopes)
#     client = gspread.authorize(credentials)
#     sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1HFCXfq3XDNokyQI3F3_yi0Jpr8ie9TK16j-zOQTAAT4/edit#gid=1711565043").get_worksheet(1)
#     return sheet

def getsheet_lineid2(): # 身分證對應
    scopes = ["https://spreadsheets.google.com/feeds"]
    global PATH
    cred_path = PATH
    credentials = ServiceAccountCredentials.from_json_keyfile_name(cred_path, scopes)
    client = gspread.authorize(credentials)
    sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1xFUgSiqU08UqjkYeB-wv8GiVnYFeT1s1MfPySVqYhQg/edit#gid=1112453375").get_worksheet(1)
    return sheet

def insertevent(time, name, category, note):
    SCOPES = ['https://www.googleapis.com/auth/calendar']
    global PATH2
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(PATH2, SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    try:
        service = build('calendar', 'v3', credentials=creds)
        appoint_time = rfc3339.rfc3339(time)
        end_time = rfc3339.rfc3339(time + datetime.timedelta(minutes=15))
        event = {
        'summary': f'{name} {category} 預約',
        'description': note,
        'start': {'dateTime': appoint_time,'timeZone': 'Asia/Taipei',},
        'end': {'dateTime': end_time,'timeZone': 'Asia/Taipei',},
        'recurrence': ['RRULE:FREQ=DAILY;COUNT=1'],}
        service.events().insert(calendarId='primary', body=event).execute()

    except HttpError as error:
        print('An error occurred: %s' % error)

@csrf_exempt
def callback(request):
 
    if request.method == 'POST':
        signature = request.META['HTTP_X_LINE_SIGNATURE']
        body = request.body.decode('utf-8')
        try:
            events = parser.parse(body, signature)  # 傳入的事件
        except InvalidSignatureError:
            return HttpResponseForbidden()
        except LineBotApiError:
            return HttpResponseBadRequest()
 
        for event in events:
            if isinstance(event, MessageEvent):  
                if event.message.text == '線上預約':
                    reply_arr = []
                    reply_arr.append(TextSendMessage(text='請完成預約門診單\n完成後我們會請專人與您聯繫'))
                    reply_arr.append(TextSendMessage(text='https://liff.line.me/1657869644-Z9PmDXzq'))
                    line_bot_api.reply_message(event.reply_token,reply_arr)

                elif event.message.text == "基本資料填寫":
                    userid = event.source.user_id
                    line_corr = getsheet_lineid2()
                    id_number = ''
                    
                    # 根據line對應表中找相應的身分證號
                    for record in line_corr.get_all_records():
                        if userid == record['LINE PID']:
                            id_number = record['身分證字號']
                            break

                    if id_number != '':
                        message=TextSendMessage(text=f'⚠️只允許申請單筆帳號\n您目前設定之身分證號為\n\n{id_number}\n\n如需修改請洽本診所')
                        line_bot_api.reply_message(event.reply_token, message)
                    else:
                        message=TextSendMessage(text=f'📢 歡迎您申請本系統\n\n👇  請直接輸入身分證號進行驗證(請使用大寫字母)')
                        line_bot_api.reply_message(event.reply_token, message)

                elif (65 <= ord(event.message.text[0]) <= 90) and len(event.message.text) == 10:     
                    userid = event.source.user_id
                    # line_corr = getsheet_lineid1()  
                    line_corr2 = getsheet_lineid2()

                    id_number = ''
                    for record in line_corr2.get_all_records():
                        if userid == record['LINE PID']:
                            id_number = record['身分證字號']
                            break
                    
                    reply_arr = []
                    if id_number == '':
                        # line_corr.append_row([userid, event.message.text])
                        line_corr2.append_row([userid, event.message.text])
                        reply_arr.append(TextSendMessage(text='請接續輸入您的基本資料\n完成後我們會請專人與您聯繫'))
                        reply_arr.append(TextSendMessage(text='https://liff.line.me/1657869644-q9LWMR19'))
                        line_bot_api.reply_message(event.reply_token,reply_arr)
                    else:
                        reply_arr.append(TextSendMessage(text=f'⚠️只允許申請單筆帳號\n您目前設定之身分證號為\n\n{id_number}\n\n如需修改請洽本診所'))
                        line_bot_api.reply_message(event.reply_token,reply_arr)
                
                elif event.message.text == '常見問題':
                    line_bot_api.reply_message(event.reply_token,
                        TemplateSendMessage(alt_text='Carousel template',
                            template=CarouselTemplate(
                                columns=[
                                    CarouselColumn(title='服務時間', text='請選擇您的問題',
                                        actions=[
                                            PostbackAction(label='抽血服務時間',text='抽血服務時間',data='Question抽血服務時間'),
                                            PostbackAction(label='門診服務時間',text='門診服務時間',data='Question門診服務時間'),]),
                                    CarouselColumn(title='資料申請', text='請選擇您的問題',
                                        actions=[
                                            PostbackAction(label='光碟申請',text='光碟申請',data='Question光碟申請'),
                                            PostbackAction(label='病歷資料複製申請',text='病歷資料複製申請',data='Question病歷資料複製申請'),
                                            ])])))
                
                elif event.message.text == 'oi3913kkzopelakeeqw': # 群發問卷
                    all_member = pd.DataFrame(getsheet_lineid2().get_all_records())
                    line_id_list = list(all_member['LINE PID'])
                    for line_id in line_id_list:
                        line_bot_api.push_message(line_id, TextSendMessage(text='https://liff.line.me/1657869644-QWMo2zZL'))
                
                elif event.message.text == 'se9dm132edosd9e83': # 推播金鑰
                    appoint_table = pd.DataFrame(getsheet_appoint().get_all_records())
                    appoint_table = appoint_table[appoint_table['預約確認']=='Y']
                    tomorrow = (datetime.datetime.now() + datetime.timedelta(hours=56)).date()

                    notify_id = []
                    time_list = []
                    for i in range(len(appoint_table)):
                        timestamp = datetime.datetime.strptime(str(appoint_table.iloc[i]['就診日']), "%Y/%m/%d")
                        timestamp2 = datetime.datetime.strptime(str(appoint_table.iloc[i]['回診日']), "%Y/%m/%d")
                        if timestamp.date() == tomorrow:
                            notify_id.append(appoint_table.iloc[i]['身分證字號'])
                            time_list.append(appoint_table.iloc[i]['就診時間'])
                        if timestamp2.date() == tomorrow:
                            notify_id.append(appoint_table.iloc[i]['身分證字號'])
                            time_list.append(appoint_table.iloc[i]['回診時間'])

                    for i in range(len(notify_id)):
                        line_to_id = pd.DataFrame(getsheet_lineid2().get_all_records())
                        index = list(line_to_id['身分證字號']).index(notify_id[i])
                        line_bot_api.push_message(line_to_id['LINE PID'][index],
                        TextSendMessage(text=f'💬 診所提醒\n您的預約診所看診時間為{tomorrow} {time_list[i]}\n-----------------------------------------\n**目前聖路加門診可提供視訊診療服務(與醫師使用LINE視訊)看診後可線上匯款。自處方箋開立當日起三天內（不含週日）請備妥病人健保卡至輔大醫院一樓藥局領藥，若有需要改看視訊診 請跟我們聯繫。\n\
                                                \n\n門診服務電話:(週一至週五08:00~16:30) 02-8512-8900/0905-301-705 地址：新北市泰山區貴子路69號15樓LINE ID :85128900 ; Wechat ID : f85128900 輔大醫院聖路加門診。'))
                
                elif event.message.text == 'e5gf5nojeo139g10guy': # 更新金鑰
                    appoint_table = pd.DataFrame(getsheet_appoint().get_all_records())
                    # 用來判斷已經成功預約 但還沒放進日曆的活動
                    appoint_table_bool = (appoint_table['預約確認']=='Y') & (appoint_table['日曆活動']== '')
                    if sum(appoint_table_bool) != 0:
                        for i in range(len(appoint_table_bool)):
                            if appoint_table_bool[i] == True:
                                date = appoint_table.at[i, '就診日']
                                time = appoint_table.at[i, '就診時間']
                                name = appoint_table.at[i, '姓名']
                                category = appoint_table.at[i, '科別']
                                time_format = datetime.datetime.strptime((date+' '+time), "%Y/%m/%d %H:%M") + datetime.timedelta(hours=-8) 
                                insertevent(time_format, name, category, note=' ')
                                
                                sheet = getsheet_appoint()
                                sheet.update_cell(i+2, 14, 'Done')

                
            elif isinstance(event, PostbackEvent):
                
                if event.postback.data[0:8] == 'Question':
                    QA_dic = {'抽血服務時間':'您好:\n本院二樓檢驗科抽血服務時間如下：\n早上07:00 開放抽號碼牌\
                                \n●週一至週五：07:30~21:30 \n●週六：07:30~13:00\n●週日暫停服務\
                                \n目前週一與週六上午為尖峰時段 & 週一至週五夜診17:00-21:30 服務櫃位較少，為避免您的等候時間過長，可選擇其他時間前來抽血。',
                            '光碟申請':'門診病友申請流程：病友至一樓影像醫學科櫃台→承辦人核對相關證件並約時間取件→ 2F批掛櫃台繳費→憑收據至影醫櫃台取件。如非病人本人申請，需查驗 病人及 被 委託人雙證件。',
                            '病歷資料複製申請':'您好\n提供您本院病歷資料複製申請說明 提供您參考 謝謝。\
                                \n●臨櫃申請:請至2樓批價掛號櫃台填寫申請表辦理。\
                                \n●傳真申請：請將病歷資料影印申請表 文件及申請人身分證正反面 傳真至(02)8512-8897。(週一到週五8:00-17:00，傳真後可來電(02)8512-8888 轉 分機23035、23032確認傳真收件無誤。​\
                                \n●受理時間:\n週一至週五：上午8:00~21:00\n週六：上午8:00~12:00\n連結網址: https://www.hospital.fju.edu.tw/Guide?FuncID=medrecord',
                            '門診服務時間':'聖路加門診服務時間(週一至週五08:00~16:30)\n上班時間(週一至週五08:00~16:30)'}
                    question = event.postback.data[8:]
                    message=TextSendMessage(text=f'{QA_dic[question]}')
                    line_bot_api.reply_message(event.reply_token, message)

        return HttpResponse()
    else:
        return HttpResponseBadRequest()

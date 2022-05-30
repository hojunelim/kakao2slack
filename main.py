from datetime import datetime
from email import header
from lib2to3.pgen2 import token
import time, win32con, win32api, win32gui, ctypes
from wsgiref import headers
from urllib import request, response
import requests
from pywinauto import clipboard  # 채팅창내용 가져오기 위해
import pandas as pd  # 가져온 채팅내용 DF로 쓸거라서
from ctypes import windll

import random

# # 카톡창 이름, (활성화 상태의 열려있는 창)
kakao_opentalk_name = ['vvip','wife','total','ikwan','pock']

PBYTE256 = ctypes.c_ubyte * 256
_user32 = ctypes.WinDLL("user32")
GetKeyboardState = _user32.GetKeyboardState
SetKeyboardState = _user32.SetKeyboardState
PostMessage = win32api.PostMessage
SendMessage = win32gui.SendMessage
FindWindow = win32gui.FindWindow
IsWindow = win32gui.IsWindow
GetCurrentThreadId = win32api.GetCurrentThreadId
GetWindowThreadProcessId = _user32.GetWindowThreadProcessId
AttachThreadInput = _user32.AttachThreadInput

MapVirtualKeyA = _user32.MapVirtualKeyA
MapVirtualKeyW = _user32.MapVirtualKeyW

MakeLong = win32api.MAKELONG
w = win32con

# # 채팅방에 메시지 전송
def kakao_sendtext(chatroom_name, text):
    # # 핸들 _ 채팅방
    hwndMain = win32gui.FindWindow(None, chatroom_name)
    hwndEdit = win32gui.FindWindowEx(hwndMain, None, "RICHEDIT50W", None)

    win32api.SendMessage(hwndEdit, win32con.WM_SETTEXT, 0, text)
    SendReturn(hwndEdit)

# # 채팅내용 가져오기
def copy_chatroom(chatroom_name):
    # # 핸들 _ 채팅방
    hwndMain = win32gui.FindWindow(None, chatroom_name)
    hwndListControl = win32gui.FindWindowEx(hwndMain, None, "EVA_VH_ListControl_Dblclk", None)

    
        
    # #조합키, 본문을 클립보드에 복사 ( ctl + c , v )
    PostKeyEx(hwndListControl, ord('A'), [w.VK_CONTROL], False)
    time.sleep(0.5)
    PostKeyEx(hwndListControl, ord('C'), [w.VK_CONTROL], False)
    try:
        ctext = clipboard.GetData()
    except:
        ctext = ''
    # print(ctext)
    if windll.user32.OpenClipboard(None):
        windll.user32.EmptyClipboard()
        windll.user32.CloseClipboard()
    return ctext

# 조합키 쓰기 위해
def PostKeyEx(hwnd, key, shift, specialkey):
    if IsWindow(hwnd):

        ThreadId = GetWindowThreadProcessId(hwnd, None)

        lparam = MakeLong(0, MapVirtualKeyA(key, 0))
        msg_down = w.WM_KEYDOWN
        msg_up = w.WM_KEYUP

        if specialkey:
            lparam = lparam | 0x1000000

        if len(shift) > 0:
            pKeyBuffers = PBYTE256()
            pKeyBuffers_old = PBYTE256()

            SendMessage(hwnd, w.WM_ACTIVATE, w.WA_ACTIVE, 0)
            AttachThreadInput(GetCurrentThreadId(), ThreadId, True)
            GetKeyboardState(ctypes.byref(pKeyBuffers_old))

            for modkey in shift:
                if modkey == w.VK_MENU:
                    lparam = lparam | 0x20000000
                    msg_down = w.WM_SYSKEYDOWN
                    msg_up = w.WM_SYSKEYUP
                pKeyBuffers[modkey] |= 128

            SetKeyboardState(ctypes.byref(pKeyBuffers))
            time.sleep(0.01)
            PostMessage(hwnd, msg_down, key, lparam)
            time.sleep(0.01)
            PostMessage(hwnd, msg_up, key, lparam | 0xC0000000)
            time.sleep(0.01)
            SetKeyboardState(ctypes.byref(pKeyBuffers_old))
            time.sleep(0.01)
            AttachThreadInput(GetCurrentThreadId(), ThreadId, False)

        else:
            SendMessage(hwnd, msg_down, key, lparam)
            SendMessage(hwnd, msg_up, key, lparam | 0xC0000000)

# # 엔터
def SendReturn(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

# # 채팅방 열기
def open_chatroom(chatroom_name):
    # # # 채팅방 목록 검색하는 Edit (채팅방이 열려있지 않아도 전송 가능하기 위하여)
    hwndkakao = win32gui.FindWindow(None, "카카오톡")
    time.sleep(0.05)
    hwndkakao_edit1 = win32gui.FindWindowEx(hwndkakao, None, "EVA_ChildWindow", None)
    time.sleep(0.05)
    hwndkakao_edit2_1 = win32gui.FindWindowEx(hwndkakao_edit1, None, "EVA_Window", None)
    time.sleep(0.05)
    hwndkakao_edit2_2 = win32gui.FindWindowEx(hwndkakao_edit1, hwndkakao_edit2_1, "EVA_Window", None)
    time.sleep(0.05)                                            # ㄴ시작핸들을 첫번째 자식 핸들(친구목록) 을 줌(hwndkakao_edit2_1)
    hwndkakao_edit3 = win32gui.FindWindowEx(hwndkakao_edit2_2, None, "Edit", None)
    time.sleep(0.05)
    # # Edit에 검색 _ 입력되어있는 텍스트가 있어도 덮어쓰기됨
    win32api.SendMessage(hwndkakao_edit3, win32con.WM_SETTEXT, 0, chatroom_name)
    time.sleep(0.5)  # 안정성 위해 필요
    SendReturn(hwndkakao_edit3)
    time.sleep(0.5)
# # 채팅내용 초기 저장 _ 마지막 채팅
def chat_last_save(talk_name):
    open_chatroom(talk_name)  # 채팅방 열기
    ttext = copy_chatroom(talk_name)  # 채팅내용 가져오기
   
    a = ttext.split('\r\n')  # \r\n 으로 스플릿 __ 대화내용 인용의 경우 \r 때문에 해당안됨
    df = pd.DataFrame(a)  # DF 으로 바꾸기

    #df[0] = df[0].str.replace('\[([\S\s]+)\] \[(오전|오후)([0-9:\s]+)\] ', '', regex=True)  # 정규식으로 채팅내용만 남기기
    try:
        return df.iloc[-2, 0]
    except:
        return df

def send_slack(talk_name,msg):
    token="xoxb-3501213867331-3503594111188-l5POGSS3QdyuJ9F4fM2nLjgD"
    attachments={
        "color":"#ff0000",
        "author_name":"kakao_"+talk_name,
    }
    response =requests.post("https://slack.com/api/chat.postMessage",
                headers={"Authorization":"Bearer "+token},
                data={
                    "channel":"#kakao_"+talk_name,
                    "text":msg,
                    "attachments":attachments,
                }
    )

# # 채팅방 커멘드 체크
def chat_chek_command(talk_name,clst):
    last_value = chat_last_save(talk_name)
    try:
        if last_value == clst:
            print("no message"+ str(datetime.now()))
            return last_value
        else:
            print("new message!!"+ str(datetime.now()))
            df1 = last_value  # 최근 채팅내용만 남김
            df2 = df1.split(' ')[0]
            save_data = df1.replace(df2, "")
            save_data = save_data.replace(" ", "")
            
            print(talk_name+" : "+df1)
        # if "@" in df2:
            send_slack(talk_name,df1)
            
            return last_value
    except:
        return last_value

def main():
    clst={}
    for talk_name in kakao_opentalk_name:
        clst[talk_name] = chat_last_save(talk_name)  # 초기설정 _ 마지막채팅 저장
        time.sleep(0.5)

    while True:
         for talk_name in kakao_opentalk_name:
            clst[talk_name] = chat_chek_command(talk_name,clst[talk_name]) # 커맨드 체크

if __name__ == '__main__':
    main()
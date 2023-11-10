from pywinauto.application import Application
from pywinauto.timings import  TimeoutError
from pywinauto.findwindows import ElementNotFoundError
from pywinauto import findwindows
import pyautogui
import pyperclip
import time

TEAMVIEWER_PATH = r"Your TeamViewer File Path"
TEAMVIEWER_TITLE = "TeamViewer"
TEAMVIEWER_IS_USED = False

session_code = "Session ID"
run_time = "Limit Time"

# 紀錄執行時間
teamviewer_app = ""
teamviewer_app_remote = ""
start_time = time.time()

def open_teamviewer():
    global TEAMVIEWER_IS_USED
    global teamviewer_app
    teamviewer_app = Application(backend="uia").start(TEAMVIEWER_PATH)
    
    time.sleep(1)
    teamviewer_app = Application(backend="uia").connect(title=TEAMVIEWER_TITLE)

    time.sleep(1)
    main_window = teamviewer_app.window(title=TEAMVIEWER_TITLE)

    time.sleep(1)
    main_window.set_focus()
    remote_support_button = main_window.child_window(title="Remote Support", control_type="Button")
    remote_support_button.click()
    print("點擊Remote Support")

    time.sleep(1)
    join_session_button = main_window.child_window(title="Join a session", control_type="Button")
    join_session_button.click()
    print("點擊Join a session")

    time.sleep(1)
    pyperclip.copy(session_code)
    pyautogui.hotkey('ctrl', 'v')
    print("輸入SessionID")

    time.sleep(1)
    connect_button = main_window.child_window(title="Connect", control_type="Button")
    connect_button.click()
    print("點擊連線")
    TEAMVIEWER_IS_USED =True

def open_waiting_room() :
    try :
        time.sleep(1)
        waiting_room_app = Application().connect(title_re=".*TeamViewer.*Waiting room.*")
        print("連接 Waiting Room")

        # 等待視窗
        time.sleep(1)
        waiting_window_app = Application(backend="uia").connect(title_re=".*TeamViewer.*Waiting room.*")
        waiting_window = waiting_window_app.window(title_re="TeamViewer - Waiting room")

        click_sure_connect_button = waiting_window.child_window(title="加入工作階段", control_type="Button").wait('ready', timeout=10)
        click_sure_connect_button.click()
        print("點擊加入工作階段")

        return waiting_room_app

    except Exception as e:
        print(f"操作人員尚未加入Session：{e}")
        return None

def check_teamviewer_status():
        global teamviewer_app
        time.sleep(1)
        main_window = teamviewer_app.window(title=TEAMVIEWER_TITLE)
        status_texts = ["Waiting for user to join", "Ready", "Ongoing"]
        for status in status_texts:
            try:
                if main_window.child_window(title=status, control_type="Text").exists():
                    print(f"{status}。")
                    return status
            except ElementNotFoundError:
                continue
        print("TeamViewer不在远程会话中。")
        return "not_connected"

def check_teamviewer_is_contect():
    try:
        global teamviewer_app_remote
        handle = findwindows.find_window(title_re="TeamViewer面板")

        app = Application().connect(handle=handle)
        window = app.window(handle=handle)
        teamviewer_app_remote = window
        if window.exists():
            print("TeamViewer 連線中")
            return True
        else :
            print("TeamViewer 未連線中")
            return False

    except IndexError:  
        print("TeamViewer未在連線。")
        return False
    except Exception as e:
        print(f"檢查TeamViewer連線狀態時發生錯誤：{e}")
        return False
    
def check_contect_time ():
    current_time = time.time()
    if( current_time - start_time ) > run_time * 60 :
        window =  teamviewer_app.window(title=TEAMVIEWER_TITLE)
        window.close()
        teamviewer_app_remote.close()
        print("預約時間已到期")
        return True


# 第一次開啟TeamViewer
while True:
    is_timeout = check_contect_time()
    if is_timeout :
        break
    try :
        open_teamviewer()
        break
    except Exception as er :
        print(f"TeamViewer開啟失敗 : {er}")
        open_teamviewer()
        if TEAMVIEWER_IS_USED :
            break

# 預約期間內保持TeamViewer 隨時可被連線控制
while True :
    # 計算預約時間是否到期
    is_timeout = check_contect_time()
    if is_timeout :
        break
        
    # TeamViewer 開啟狀態開始監控Waiting Room
    if TEAMVIEWER_IS_USED  :
        try:
            while True:
                try:
                    # 監聽 Waiting room                     
                    status = check_teamviewer_status()
                    # TeamView 縮小後無法透過 check_teamviewer_status 判斷狀態，連線後改用TeamViewer 面板監控是否被連線
                    is_contect = check_teamviewer_is_contect()

                    # 
                    if status == "Ready" and not is_contect:
                        open_waiting_room()
                        time.sleep(2)
                        is_contect = check_teamviewer_is_contect()

                    # 連線中斷後，返回Waiting Room 備便使用者下次連線
                    if not is_contect  :
                        open_teamviewer()
                        open_waiting_room()

                    current_time = time.time()
                    is_timeout = check_contect_time()
                    if is_timeout :
                        break

                # 監聽Waiting Room 發生問題或因為沒有人連線導致Waiting Room自動關閉，重新開啟 Waiting Room 
                except Exception as e:
                    print(f"Waiting Room Error：{e}")
                    open_teamviewer()

        except Exception as e:
            print(f"Error Msg：{e}")



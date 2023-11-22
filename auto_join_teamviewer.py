import sys
import io
import os
import time
from datetime import datetime
from pywinauto.application import Application
from pywinauto.findwindows import ElementNotFoundError
from pywinauto import findwindows
import pyautogui
import pyperclip
import psutil
from logger import Logger

logger = Logger().get_logger()
# setting utf-8 for net console
os.environ["PYTHONUTF8"] = "1"
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# time style
def record_time():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

logger.info(f"{record_time()} : exe start")

# Input params
if len(sys.argv) < 4:
    logger.info("Error Msg : Please offer code and runnimg time")
    sys.exit(1)

session_code = sys.argv[1]
run_time = int(sys.argv[2])
full_path_to_exe = sys.argv[3]

# session_code = 149572100
# run_time = 3
# full_path_to_exe = r"D:\Intersense\ADX\no_limit_publish\TeamViewerAutoConnect.exe"

logger.info(f"Input Params Code : {session_code} , RunningTime : {run_time} , FilePath : {full_path_to_exe}")

TEAMVIEWER_PATH = r"C:\Program Files\TeamViewer\TeamViewer.exe"
TEAMVIEWER_TITLE = "TeamViewer"
TEAMVIEWER_IS_USED = False

# app
teamviewer_app = ""
teamviewer_app_remote = ""
# time
start_time = time.time()

def cancel_teamviewer_app():
    for proc in psutil.process_iter():
        try:
            process_name = proc.name()
            executable_path = proc.exe()
            if "TeamViewer" in process_name and "TeamViewerAutoConnect.exe" not in executable_path:
                proc.terminate()
                proc.wait()  
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass 

# Open TeamViewer app
def open_teamviewer():
    global TEAMVIEWER_IS_USED
    global teamviewer_app
    teamviewer_app = Application(backend="uia").start(TEAMVIEWER_PATH)
    try :
        time.sleep(4)
        teamviewer_app = Application(backend="uia").connect(title=TEAMVIEWER_TITLE)

        time.sleep(1)
        main_window = teamviewer_app.window(title=TEAMVIEWER_TITLE)

        time.sleep(1)
        main_window.set_focus()
        remote_support_button = main_window.child_window(title="Remote Support", control_type="Button")
        remote_support_button.click()
        print("Clieck Remote Support")

        time.sleep(1)
        join_session_button = main_window.child_window(title="Join a session", control_type="Button")
        join_session_button.click()
        print("Clieck Join a session")
        main_window.set_focus()
        time.sleep(1)
        pyperclip.copy(session_code)
        pyautogui.hotkey('ctrl', 'v')
        print("Input SessionID")

        time.sleep(1)
        connect_button = main_window.child_window(title="Connect", control_type="Button")
        connect_button.click()
        print("Click Conntect")
        TEAMVIEWER_IS_USED = True
    
    except Exception as err :
        logger.error(err)
        cancel_teamviewer_app()
        TEAMVIEWER_IS_USED = False


# Waiting Supporter Connect  
def open_waiting_room() :
    try :
        # Waitting Room 
        time.sleep(1)
        print("Waiting Room")
        waiting_window_app = Application(backend="uia").connect(title_re=".*TeamViewer.*Waiting room.*")
        time.sleep(1)
        waiting_window = waiting_window_app.window(title_re="TeamViewer - Waiting room")
        waiting_window.set_focus()
        print("Search click join worj btn")
        click_sure_connect_button = waiting_window.child_window(title="Join session", control_type="Button").wait('ready', timeout= 30)
        click_sure_connect_button.click()
        print("click join work")

    except Exception as e:
        if e == "timed out" :
            logger.info("No Person Join")
        else :
            logger.error(e)

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
        print("not join session")
        return "not_connected"

def check_teamviewer_is_contect():
    try:
        global teamviewer_app_remote
        handle = findwindows.find_window(title_re="TeamViewer Panel")

        app = Application().connect(handle=handle)
        window = app.window(handle=handle)
        teamviewer_app_remote = window
        if window.exists():
            print("TeamViewer Connect")
            return True
        else :
            print("TeamViewer not Connect")
            return False

    except IndexError:  
        print("TeamViewer not Connect")
        return False
    except Exception as err:
        print(f"Check TeamViewer Connect Error：{err}")
        return False

def check_contect_time ():
    current_time = time.time()
    if( current_time - start_time ) > run_time * 60 :
        close_connect_teamviewer()
        logger.info(f"{record_time()} : Timeout and Exe is end")
        sys.exit(0)

# Cancel Connect
def close_connect_teamviewer():
    # Close TeamViewer Panel
    try :
        handle = findwindows.find_window(title_re="TeamViewer Panel")
        if handle:
            panel_app = Application().connect(handle=handle)
            panel_window = panel_app.window(handle=handle)
            panel_window.close()
    except Exception as err :
        logger.info(f"{record_time()} TeamViewer Panel   Not Find")
    
    # Close TeamViewer Waiting Room 
    try :
        waiting_window_app = Application(backend="uia").connect(title_re=".*TeamViewer.*Waiting room.*")
        waiting_window = waiting_window_app.window(title_re="TeamViewer - Waiting room")
        waiting_window.close()
    except Exception as err : 
        logger.info(f"{record_time()} : TeamViewer - Waiting room  Not Find")

    # Close TeamViwer Window
    try :
        teamviewer_app = Application(backend="uia").start(TEAMVIEWER_PATH)
        time.sleep(2)
        teamviewer_app = Application(backend="uia").connect(title=TEAMVIEWER_TITLE)
        time.sleep(1)
        main_window = teamviewer_app.window(title=TEAMVIEWER_TITLE)
        main_window.close()
    except Exception as err : 
        logger.info(f"{record_time()} TeamViewer Not Find")

    # Check autoJoinTeamViewer.exe is running 
    if not is_exe_running(full_path_to_exe):
        # open exe
        os.startfile(full_path_to_exe)
        print(f"Started {full_path_to_exe}")

def is_exe_running(exe_path):
    for process in psutil.process_iter(['pid', 'exe']):
        if process.info['exe'] == exe_path:
            return True
    return False

# Firest Open TeamViewer
while True:
    is_timeout = check_contect_time()
    try :
        open_teamviewer()
        if TEAMVIEWER_IS_USED :
            break
    except Exception as er :
        print(f"TeamViewer Open Error Msg : {er}")
        open_teamviewer()
        if TEAMVIEWER_IS_USED :
            break

# Waiting Join
while True:
    try:
        # check Waiting room                     
        status = check_teamviewer_status()
        time.sleep(1)

        # After minimizing TeamViewer, it is not possible to determine the status through check_teamviewer_status. Once connected, 
        # switch to monitoring the connection status using the TeamViewer panel
        is_contect = check_teamviewer_is_contect()
        time.sleep(1)

        if status == "Ready" and not is_contect:
            open_waiting_room()
            time.sleep(3)
            is_contect = check_teamviewer_is_contect()

        # Connect Dispose，Return Waiting Room 
        if not is_contect  :
            open_teamviewer()
            time.sleep(1)
        
        is_timeout = check_contect_time()
        time.sleep(1)
    # If there are issues monitoring the Waiting Room or it automatically closes due to no one connecting, reopen the Waiting Room.
    except Exception as e:
        logger.error(f"Waiting Room Error：{e}")
        # Close Finish Before Open New TeamView
        cancel_teamviewer_app()
        # for proc in psutil.process_iter():
        #     try:
        #         process_name = proc.name()
        #         executable_path = proc.exe()
        #         if "TeamViewer" in process_name and "TeamViewerAutoConnect.exe" not in executable_path:
        #             proc.terminate()
        #             proc.wait()  
        #     except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
        #         pass 
        #     except Exception as err :
        #         logger.error(f"repeat err msg : {err}")
        open_teamviewer()



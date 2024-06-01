from pywinauto.application import Application
from pywinauto.findwindows import ElementNotFoundError
from pywinauto import findwindows
from pywinauto.timings import wait_until
import pyautogui
import pyperclip
from logger import Logger
import psutil
import time

logger = Logger().get_logger()
TEAMVIEWER_PATH = r"C:\Program Files\TeamViewer\TeamViewer.exe"
TEAMVIEWER_TITLE = "TeamViewer"
pyautogui.FAILSAFE = False

class TeamViewerOperations:
    @staticmethod
    def start_teamviewer():
        app = Application(backend="uia").start(TEAMVIEWER_PATH)
        return app
    
    @staticmethod
    def connect_to_teamviewer():
        app = Application(backend="uia").connect(title=TEAMVIEWER_TITLE)
        return app
    
    @staticmethod
    def get_main_window(app):
        return app.window(title= TEAMVIEWER_TITLE)
    
    @staticmethod
    def click_remote_support(main_window):
        print("Clieck Remote Support")
        main_window.set_focus()
        remote_support_button = main_window.child_window(title="Remote Support", control_type="Button")
        remote_support_button.click()

    @staticmethod
    def click_join_session(main_window):
        print("Clieck Join a session")
        join_session_button = main_window.child_window(title="Join a session", control_type="Button")
        join_session_button.click()

    @staticmethod
    def copy_session_code(session_code) :
        pyperclip.copy(session_code)
        pyautogui.hotkey('ctrl', 'v')
        print("Input SessionID")

    @staticmethod
    def accept_join_remote(main_window):
        main_window.set_focus()
        connect_button = main_window.child_window(title="Connect", control_type="Button")
        connect_button.click()
        print("Click Conntect")  

    @staticmethod
    def allow_access():
        temp_app = Application(backend="uia").connect(title_re=".*TeamViewer.*Waiting room.*")
        temp_room_window = temp_app.window(title_re=".*TeamViewer.*Waiting room.*", control_type="Window")

        temp_room_window.set_focus()

        allow_button = temp_room_window.child_window(title="Allow access", control_type="Text")
        
        # 如果按鈕不存在，可能需要更具體的查找
        if not allow_button.exists():
            allow_button = temp_room_window.child_window(control_type="Text", found_index=0)
            if allow_button.window_text() == "Allow access":
                print("使用不同方法找到 Allow access 按鈕")
                
        allow_button.click_input()
        print("點擊 Allow Access")

    @staticmethod
    def close_window() :
        teamviewer_app = Application(backend="uia").start(TEAMVIEWER_PATH )
        time.sleep(2)
        teamviewer_app = Application(backend="uia").connect(title=TEAMVIEWER_TITLE)
        time.sleep(1)
        main_window = teamviewer_app.window(title=TEAMVIEWER_TITLE)
        main_window.close()



class TeamViewerWaitingRoomOperations:
    @staticmethod
    def connect_to_waiting_room():
        print("Connect Waiting Room")
        app = Application(backend="uia").connect(title_re=".*TeamViewer.*Waiting room.*")
        return app
    
    @staticmethod
    def get_waiting_romm_window(waiting_window_app):
        app = waiting_window_app = waiting_window_app.window(title_re="TeamViewer - Waiting room")
        return app
    
    @staticmethod
    def waiting_for_supporter_join(waiting_window):
        try :
            print("Wait For Supporter Join")
            btn = waiting_window.child_window(title="Join session", control_type="Button").wait('ready', timeout= 15)
            waiting_window.set_focus()
            btn.click()
            print("click join work")
        except Exception as e:
            print(e)
            logger.info(f"Wait For Supporter Join {e}")
   
    @staticmethod
    def close_teamviewer_waiting_room():
        waiting_window_app = Application(backend="uia").connect(title_re=".*TeamViewer.*Waiting room.*")
        waiting_window = waiting_window_app.window(title_re="TeamViewer - Waiting room")
        waiting_window.close()

class CancelTeamViewerExe :
    @staticmethod
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

class CheckTeamViewerStatus :
    @staticmethod
    def check_teamviewer_status(teamviewer_app):
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
    
    @staticmethod
    def check_teamviewer_is_contect(teamviewer_app_remote) :
        teamviewer_app_remote
        if teamviewer_app_remote.exists():
            print("TeamViewer Connect")
            return True
        else :
            print("TeamViewer not Connect")
            return False
        
    @staticmethod
    def check_contect_time(start_time,run_time) :
        current_time = time.time()
        if( current_time - start_time ) > run_time * 60 :
            return True
        else :
            return False
        
class TeamViewerPanel():
    @staticmethod
    def get_teamviewer_panel_window():
        handle = findwindows.find_window(title_re="TeamViewer Panel")
        app = Application().connect(handle=handle)
        window = app.window(handle=handle)
        return window
    @staticmethod
    def close_teamviewer_Panel():
        handle = findwindows.find_window(title_re="TeamViewer Panel")
        panel_app = Application().connect(handle=handle)
        panel_window = panel_app.window(handle=handle)
        panel_window.close()
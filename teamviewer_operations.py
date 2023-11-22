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

class TeamViewerOperatiomns:
    def start_teamviewer(teamviewer_path):
        app = Application(backend="uia").start(teamviewer_path)
        return app

    def connect_to_teamviewer(app_title):
        app = Application(backend="uia").connect(title=app_title)
        return app

    def get_main_window(app,app_title):
        return app.window(title= app_title)

    def click_remote_support(main_window):
        print("Clieck Remote Support")
        main_window.set_focus()
        remote_support_button = main_window.child_window(title="Remote Support", control_type="Button")
        remote_support_button.click()

    
    def click_join_session(main_window):
        print("Clieck Join a session")
        join_session_button = main_window.child_window(title="Join a session", control_type="Button")
        join_session_button.click()

    def copy_session_code(session_code) :
        pyperclip.copy(session_code)
        pyautogui.hotkey('ctrl', 'v')
        print("Input SessionID")

    def accept_join_remote(main_window):
        connect_button = main_window.child_window(title="Connect", control_type="Button")
        connect_button.click()
        print("Click Conntect")   

    def close_window(TEAMVIEWER_PATH,TEAMVIEWER_TITLE) :
        teamviewer_app = Application(backend="uia").start(TEAMVIEWER_PATH)
        time.sleep(2)
        teamviewer_app = Application(backend="uia").connect(title=TEAMVIEWER_TITLE)
        time.sleep(1)
        main_window = teamviewer_app.window(title=TEAMVIEWER_TITLE)
        main_window.close()

class TeamViewerWaitingRoomOperations:
    def connect_to_waiting_room():
        print("Connect Waiting Room")
        app = Application(backend="uia").connect(title_re=".*TeamViewer.*Waiting room.*")
        return app

    def get_waiting_romm_window(waiting_window_app):
        app = waiting_window_app = waiting_window_app.window(title_re="TeamViewer - Waiting room")
        return app
    
    def waiting_for_supporter_join(waiting_window):
        try :
            print("Wait For Supporter Join")
            btn = waiting_window.child_window(title="Join session", control_type="Button").wait('ready', timeout= 30)
            btn.click()
            print("click join work")
        except Exception as e:
            if e == "timed out" :
                print("supporter not join")
                logger.info("No supporter Join")
            else :
                print(e)
                logger.error(e)
    
    def close_teamviewer_waiting_room():
        waiting_window_app = Application(backend="uia").connect(title_re=".*TeamViewer.*Waiting room.*")
        waiting_window = waiting_window_app.window(title_re="TeamViewer - Waiting room")
        waiting_window.close()

class CancelTeamViewerExe :
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

class CheckTeamViwerStatus :
    def check_teamviewer_status(teamviewer_app,TEAMVIEWER_TITLE):
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
    
    def check_teamviewer_is_contect(teamviewer_app_remote) :
        teamviewer_app_remote
        if teamviewer_app_remote.exists():
            print("TeamViewer Connect")
            return True
        else :
            print("TeamViewer not Connect")
            return False
    def check_contect_time(start_time,run_time) :
        current_time = time.time()
        if( current_time - start_time ) > run_time * 60 :
            return True
        else :
            return False
        
class TeamViewerPanel():
    def get_teamviewer_panel_window():
        handle = findwindows.find_window(title_re="TeamViewer Panel")
        app = Application().connect(handle=handle)
        window = app.window(handle=handle)
        return window
    
    def close_teamviewer_Panel():
        handle = findwindows.find_window(title_re="TeamViewer Panel")
        panel_app = Application().connect(handle=handle)
        panel_window = panel_app.window(handle=handle)
        panel_window.close()
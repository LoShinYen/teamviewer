import sys
import io
import os
import time
import psutil
from logger import Logger
import teamviewer_operations

logger = Logger().get_logger()

teamviewr_main_operation = teamviewer_operations.TeamViewerOperations()
teamviwer_waitingromm_operation = teamviewer_operations.TeamViewerWaitingRoomOperations()
teamviwer_cancel_operation = teamviewer_operations.CancelTeamViewerExe()
teamviwer_check_status = teamviewer_operations.CheckTeamViewerStatus()
teamviwer_panel_operation = teamviewer_operations.TeamViewerPanel()

# setting utf-8 for net console
os.environ["PYTHONUTF8"] = "1"
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

logger.info("exe start")

# # Input params
if len(sys.argv) < 4:
    logger.info("Error Msg : Please offer code and runnimg time")
    sys.exit(1)

session_code = sys.argv[1]
run_time = int(sys.argv[2])
full_path_to_exe = sys.argv[3]
TEAMVIEWER_IS_USED = False
# session_code = 141408465
# run_time = 3
# full_path_to_exe = r"D:\Intersense\ADX\no_limit_publish\TeamViewerAutoConnect.exe"

logger.info(f"Input Params Code : {session_code} , RunningTime : {run_time} , FilePath : {full_path_to_exe}")

# app
teamviewer_app = ""
teamviewer_app_remote = ""
# time
start_time = time.time()

# Open TeamViewer app
def open_teamviewer():
    global TEAMVIEWER_IS_USED
    global teamviewer_app
    teamviewer_app = teamviewr_main_operation.start_teamviewer()
    try :
        time.sleep(4)
        teamviewer_app = teamviewr_main_operation.connect_to_teamviewer()

        time.sleep(1)
        main_window = teamviewr_main_operation.get_main_window(teamviewer_app)

        time.sleep(1)
        teamviewr_main_operation.click_remote_support(main_window)

        time.sleep(1)
        teamviewr_main_operation.click_join_session(main_window)
        main_window.set_focus()

        time.sleep(1)
        teamviewr_main_operation.copy_session_code(session_code)

        time.sleep(1)
        teamviewr_main_operation.accept_join_remote(main_window)
        TEAMVIEWER_IS_USED = True
    
    except Exception as err :
        logger.error(err)
        teamviwer_cancel_operation.cancel_teamviewer_app()
        TEAMVIEWER_IS_USED = False

# Waiting Supporter Connect  
def open_waiting_room() :
    try :
        time.sleep(1)
        waiting_window_app = teamviwer_waitingromm_operation.connect_to_waiting_room()
        time.sleep(1)
        waiting_window = teamviwer_waitingromm_operation.get_waiting_romm_window(waiting_window_app)
        waiting_window.set_focus()
        
        teamviwer_waitingromm_operation.waiting_for_supporter_join(waiting_window)
    except Exception as err :
        logger.error(err)

def check_teamviewer_status():
    time.sleep(1)
    return teamviwer_check_status.check_teamviewer_status(teamviewer_app)

def check_teamviewer_is_contect():
    try:
        global teamviewer_app_remote
        teamviewer_app_remote = teamviwer_panel_operation.get_teamviewer_panel_window()
        return teamviwer_check_status.check_teamviewer_is_contect(teamviewer_app_remote)
    
    except IndexError:  
        print("TeamViewer not Connect")
        return False
    
    except Exception as err:
        print(f"Check TeamViewer Connect Error：{err}")
        return False

def check_contect_time ():
    if teamviwer_check_status.check_contect_time(start_time,run_time) :
        close_connect_teamviewer()
        logger.info(f"Timeout and Exe is end")
        sys.exit(0)

# Cancel Connect
def close_connect_teamviewer():
    # Close TeamViewer Panel
    try :
        teamviwer_panel_operation.close_teamviewer_Panel()
    except Exception :
        logger.info("TeamViewer Panel Not Find")
    
    # Close TeamViewer Waiting Room 
    try :
        teamviwer_waitingromm_operation.close_teamviewer_waiting_room()
    except Exception : 
        logger.info("TeamViewer - Waiting room Not Find")

    # Close TeamViwer Window
    try :
        teamviewr_main_operation.close_window()
    except Exception : 
        logger.info("TeamViewer Not Find")

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
    except Exception as err :
        print(f"TeamViewer Open Error Msg : {err}")
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
        teamviwer_cancel_operation.cancel_teamviewer_app()
        open_teamviewer()



Content : 
This Python script is primarily designed to work with the TeamViewer Web API Create Sessions, 
achieving an automated unattended access feature. The application scope can be expanded to accommodate 
the API Create or Cancel Session to manage the control status and number of Supporter personnel.

Future Expectation:
To achieve automation in obtaining the API Key through TeamViewer OAuth2 authentication.

TeamViewer Desktop Version: 15.47.3 (d11df019d35)

version : python3.12.0

Reference Package :
pip install pywinauto
pip install pyautogui
pip install pyperclip

---------------------------
pip install pyinstaller

set PYTHONUTF8=1
pyinstaller main.py
---------------------------
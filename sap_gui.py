import os
import subprocess
import time
import win32com.client
from dotenv import load_dotenv
import config

load_dotenv()

# SAP_USERNAME = os.getenv("SAP_USERNAME")
# SAP_PASSWORD = os.getenv("SAP_PASSWORD")
# SAP_CLIENT   = os.getenv("SAP_CLIENT")
# SAP_LANGUAGE = os.getenv("SAP_LANGUAGE")
# SAP_GUI_PATH         = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
# SAP_CONNECTION_NAME  = "DEMO 2022"   
SAP_USERNAME        = config.SAP_USERNAME
SAP_PASSWORD        = config.SAP_PASSWORD
SAP_CLIENT          = config.SAP_CLIENT
SAP_LANGUAGE        = config.SAP_LANGUAGE
SAP_GUI_PATH        = config.SAP_GUI_PATH
SAP_CONNECTION_NAME = config.SAP_CONNECTION_NAME 

def open_and_login():
    """Open SAP GUI, connect to DEMO 2022, log in and return the session object."""

    print("Opening SAP GUI...")
    subprocess.Popen([SAP_GUI_PATH])
    time.sleep(2)  
    

    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine

    print(f"Connecting to {SAP_CONNECTION_NAME}...")
    connection = application.OpenConnection(SAP_CONNECTION_NAME, True)
    
    session = None
    time.sleep(2)  # wait 7 seconds
    session = connection.Children(0)
    

    print("Successfully opened DEMO 2022!")
    print("Connection established!")

    session.findById("wnd[0]/usr/txtRSYST-MANDT").text = SAP_CLIENT
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = SAP_USERNAME
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = SAP_PASSWORD
    session.findById("wnd[0]/usr/txtRSYST-LANGU").text = SAP_LANGUAGE
    session.findById("wnd[0]").sendVKey(0)

    print("Logged in successfully.")
    return session

if __name__ == "__main__":
    open_and_login()

import win32com.client
import time

def get_session():
    """
    Attach to an active SAP GUI session.
    Requires that sap_gui.py has already opened SAP and logged in.
    """
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
    except Exception:
        raise Exception("SAP GUI is not running or scripting is disabled.")

    try:
        application = SapGuiAuto.GetScriptingEngine
    except Exception:
        raise Exception("Scripting engine not available. Enable SAP GUI scripting.")

    # Get first connection
    try:
        connection = application.Children(0)
    except Exception:
        raise Exception("No SAP connections found.")

    # Wait for session to be created
    for _ in range(20):
        if connection.Children.Count > 0:
            return connection.Children(0)
        time.sleep(0.5)

    raise Exception("No SAP session found. Did you log in?")

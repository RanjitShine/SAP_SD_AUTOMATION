from sap_gui import open_and_login
from sap_session import get_session
import time
import pandas as pd  # Added for Excel reading
import win32com.client
import time
import os
import sys

def resource_path(*relative_parts):

    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, *relative_parts)

def goto_tcode(session):
 
    session.findById("wnd[0]/tbar[0]/okcd").text = "/NOVXD"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[1]/btn[5]").press()
    session.findById("wnd[0]/usr/txtV_TVST-VSTEL").text
    session.findById("wnd[0]/usr/txtV_TVST-VTEXT").text
   
def Shipping_Point(session):

    df = pd.read_csv("test/Shipping_Point.csv")
    df = df.reset_index(drop=True)
   
    for index, row in df.iterrows():
        Shipping_Point = str(row["Shipping Point"]).strip()
        Description = str(row["Description"]).strip()
        country = str(row["country"]).strip()
       
        session.findById(f"wnd[0]/usr/txtV_TVST-VSTEL").text = Shipping_Point
        session.findById(f"wnd[0]/usr/txtV_TVST-VTEXT").text = Description
       
        # session.findById("wnd[0]/usr/txtV_TVST-VTEXT").setFocus
        # session.findById("wnd[0]/usr/txtV_TVST-VTEXT").caretPosition = 8
        session.findById("wnd[0]/tbar[1]/btn[28]").press()
        session.findById(f"wnd[1]/usr/ctxtADDR1_DATA-COUNTRY").text = country
        # session.findById("wnd[1]/usr/ctxtADDR1_DATA-COUNTRY").setFocus()
        # session.findById("wnd[1]/usr/ctxtADDR1_DATA-COUNTRY").caretPosition = 2
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        # session.findById("wnd[0]/tbar[0]/btn[5]").press() 
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
       
        # time.sleep(3)
        session.findById("wnd[0]/tbar[1]/btn[25]").press()
        session.findById("wnd[0]/tbar[1]/btn[5]").press()
       
def main():
     
     session = get_session()
     goto_tcode(session) 
     Shipping_Point(session) 

if __name__ == "__main__":
    main()
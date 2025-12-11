from sap_session import get_session
import time
import pandas as pd  # Added for Excel reading
import win32com.client
import os
import sys

def resource_path(*relative_parts):

    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, *relative_parts)

def goto_tcode(session):
 
    session.findById("wnd[0]/tbar[0]/okcd").text = "/NSIMG"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()
   
    session = win32com.client.GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)
   
    # Expand tree nodes
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").expandNode("02  1      6")
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").topNode = "01  1      1"

    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").expandNode("03  2      1")
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").topNode = "01  1      1"
 
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").expandNode("04  3      4")
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").topNode = ("01  1      1")
    
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").selectItem("05  4      4","2")
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").ensureVisibleHorizontalItem("05  4      4","2")
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").clickLink("05  4      4","2")
  
    session.findById("wnd[0]/tbar[1]/btn[5]").press()
    time.sleep(2)
 
 
#function to add Division.csv
def Sales_group(session):
    
    df = pd.read_csv("test/Sales_Group.csv")  
    df = df.reset_index(drop=True)  

    # df = pd.read_csv(csv_path, encoding="utf-8-sig")
    # df = df.reset_index(drop=True)
   
    for index, row in df.iterrows():
        Sales_Group = str(row["Sales Group"]).strip()
        Description = str(row["Description"]).strip()
       
        session.findById(f"wnd[0]/usr/tblSAPL0ORGCORETCTRL_V_TVKGR/txtV_TVKGR-VKGRP[0,{index}]").text = Sales_Group
        session.findById(f"wnd[0]/usr/tblSAPL0ORGCORETCTRL_V_TVKGR/txtV_TVKGR-BEZEI[1,{index}]").text = Description
 
    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

def main():
     
     session = get_session()
     goto_tcode(session) 
     Sales_group(session) 

if __name__ == "__main__":
         main()

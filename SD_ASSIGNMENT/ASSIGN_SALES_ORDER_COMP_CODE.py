from sap_gui import open_and_login
from sap_session import get_session
import pandas as pd
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
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").expandNode ("02  1      6")
            session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").topNode = "01  1      1"
            session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").expandNode ("03  2      2")
            session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").topNode = "01  1      1"
            session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").expandNode ("04  3      4")
            session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").topNode = "01  1      1"
            session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").selectItem ("05  4      1","2")
            session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").ensureVisibleHorizontalItem ("05  4      1","2")
            session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").clickLink ("05  4      1","2")
            session.findById("wnd[0]/tbar[1]/btn[5]").press()
 
  
 
def Salessales_div(session):
    csv_path = resource_path("CSV/Assign_Sales_org_and_comp_code.csv")
    df = pd.read_csv(csv_path, encoding="utf-8-sig")
    df = df.reset_index(drop=True)
# Move table to top before writing
 
    for idx, row in df.iterrows():
   
        Sales_Office = str(row["Sales_Office"]).strip()
        Sales_grp= str(row["Sales_grp"]).strip()
        session.findById(f"wnd[0]/usr/tblSAPLCUST_ASSIGN_SDTCTRL_V_TVBVK_ASSIGN/ctxtV_TVBVK_ASSIGN-VKBUR[0,{idx}]").text = Sales_Office
        session.findById(f"wnd[0]/usr/tblSAPLCUST_ASSIGN_SDTCTRL_V_TVBVK_ASSIGN/ctxtV_TVBVK_ASSIGN-VKGRP[2,{idx}]").text = Sales_grp
        session.findById("wnd[0]").sendVKey(0)
 
def main():
    open_and_login()
    session = get_session()
    goto_tcode(session)
    Salessales_div(session)
if __name__ == "__main__":
        main()
 
 
 
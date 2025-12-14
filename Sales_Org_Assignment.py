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
   
    


    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").expandNode ("02  1      6")
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").topNode = "01  1      1"
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").expandNode ("03  2      2")
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").topNode = "01  1      1"
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").expandNode ("04  3      4")
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").topNode = "01  1      1"
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").selectItem  ("05  4      2","2")
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").ensureVisibleHorizontalItem ("05  4      2","2")
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").clickLink ("05  4      2","2")
    session.findById("wnd[0]/tbar[1]/btn[5]").press()


def Sales_Assign(session):
    csv_path = resource_path("CSV/Assig_of_Sales_org_to_Dis_ch.csv")
    

    

    df = pd.read_csv(csv_path, encoding="utf-8-sig")
    df = df.reset_index(drop=True)
# Move table to top before writing

    for idx, row in df.iterrows():
        vkorg = str(row["Sales Organization"]).strip()
        vgorg = str(row["Distribution Channel"]).strip()
        session.findById(f"wnd[0]/usr/tblSAPLCUST_ASSIGN_SDTCTRL_V_TVTA_ASSIGN/ctxtV_TVTA_ASSIGN-VKORG[0,{idx}]").text = vkorg
        session.findById(f"wnd[0]/usr/tblSAPLCUST_ASSIGN_SDTCTRL_V_TVKOV_ASSIGN/ctxtV_TVKOV_ASSIGN-VTWEG[2,{idx}]").text = vgorg
        session.findById("wnd[0]/usr/tblSAPLCUST_ASSIGN_SDTCTRL_V_TVKOV_ASSIGN/ctxtV_TVKOV_ASSIGN-VTWEG[2,0]").setFocus
        session.findById("wnd[0]/usr/tblSAPLCUST_ASSIGN_SDTCTRL_V_TVKOV_ASSIGN/ctxtV_TVKOV_ASSIGN-VTWEG[2,0]").caretPosition = 1
        session.findById("wnd[0]").sendVKey(0)
    
    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()


def main():
    open_and_login() 
    session = get_session()
    goto_tcode(session) 
    Sales_Assign(session)
    

if __name__ == "__main__":
        main()

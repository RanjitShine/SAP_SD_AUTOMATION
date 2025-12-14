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
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").selectItem ("05  4      4","2")
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").ensureVisibleHorizontalItem ("05  4      4","2")
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").clickLink ("05  4      4","2")
    session.findById("wnd[0]/tbar[1]/btn[5]").press()


def Assign_Sales_Area(session):
    csv_path = resource_path("CSV/Sales_Area.csv")
    

    

    df = pd.read_csv(csv_path, encoding="utf-8-sig")


    for idx, row in df.iterrows():
        Sales_Organization = str(row["Sales Organization"]).strip()
        Distribution_Channel = str(row["Distribution Channel"]).strip()
        Division = str(row["Division"]).strip()
        session.findById(f"wnd[0]/usr/tblSAPLCUST_ASSIGN_SDTCTRL_V_TVTA_ASSIGN/ctxtV_TVTA_ASSIGN-VKORG[0,{idx}]").text = Sales_Organization
        
        

        session.findById(f"wnd[0]/usr/tblSAPLCUST_ASSIGN_SDTCTRL_V_TVTA_ASSIGN/ctxtV_TVTA_ASSIGN-VTWEG[2,{idx}]").text = Distribution_Channel
        

        session.findById(f"wnd[0]/usr/tblSAPLCUST_ASSIGN_SDTCTRL_V_TVTA_ASSIGN/ctxtV_TVTA_ASSIGN-SPART[4,{idx}]").text = Division
        session.findById("wnd[0]").sendVKey(0)


        
    
    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()



def main():
    open_and_login() 
    session = get_session()
    goto_tcode(session) 
    Assign_Sales_Area(session)
    

if __name__ == "__main__":
        main()

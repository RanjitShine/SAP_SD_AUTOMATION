from sap_gui import open_and_login
from sap_session import get_session
import time
import pandas as pd 
import sys
import os
 
def resource_path(*relative_parts):

    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, *relative_parts)

def goto_tcode(session):
   
    
    session.findById("wnd[0]/tbar[0]/okcd").text = "/NSIMG"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()
   
    # Expand tree nodes
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").expandNode("02  1      6")
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").topNode = "01  1      1"
 
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").expandNode("03  2      1")
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").topNode = "01  1      1"
 
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").expandNode("04  3      4")
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").topNode = ("01  1      1")
 
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").selectItem("05  4      3", "2")
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").ensureVisibleHorizontalItem("05  4      3", "2")
    session.findById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").clickLink("05  4      3", "2")
 
    session.findById("wnd[0]/tbar[1]/btn[5]").press()
 
def sales_office(session):
    df = pd.read_csv(r"test\Sales_office.csv")
    df = df.reset_index(drop=True)

    for index, row in df.iterrows():
        Sales_Office = str(row["Sales Office"]).strip()
        Description  = str(row["Description"]).strip()
        country      = str(row["country"]).strip()

    
        session.findById(f"wnd[0]/usr/tblSAPL0ORGTCTRL_V_TVBUR/txtV_TVBUR-VKBUR[0,{index}]").text = Sales_Office
        session.findById(f"wnd[0]/usr/tblSAPL0ORGTCTRL_V_TVBUR/txtV_TVBUR-BEZEI[2,{index}]").text = Description
        session.findById("wnd[0]/usr/tblSAPL0ORGTCTRL_V_TVBUR").getAbsoluteRow(index)

        # Open & fill popup for THIS row
        try:
            session.findById("wnd[0]/tbar[0]/btn[11]").press()  # e.g. Address button
            time.sleep(0.5)

            session.findById("wnd[1]/usr/ctxtADDR1_DATA-COUNTRY").text = country
            session.findById("wnd[1]/tbar[0]/btn[0]").press()   # OK

            try:
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except:
                pass

            print("Row", index, "country set to", country)
        except Exception as e:
            print("Popup failed for row", index, ":", e)


def main():
     
     session = get_session()
     goto_tcode(session) 
     sales_office(session) 

if __name__ == "__main__":
         main()

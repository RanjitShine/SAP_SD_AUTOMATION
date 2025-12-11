# from sap_gui import open_and_login
from sap_session import get_session
import time
import pandas as pd  # Added for Excel reading
import os
import sys


def resource_path(*relative_parts):

    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, *relative_parts)

def goto_tcode(session):
   
    # session = get_session()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nV_TVTW"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[1]/btn[5]").press()
    time.sleep(2)

def distribution_channel(session):
    # session = get_session()

    df = pd.read_csv("test/Distribution_channel.csv")  
    df = df.reset_index(drop=True)  
 
    for idx, row in df.iterrows():
        row_index=idx+1
        distribution_channel = str(row["Distribution Channel"]).strip()
        Description = str(row["Description"]).strip()

        session.findById("wnd[0]/usr/tblSAPL0ORGCORETCTRL_V_TVTW").getAbsoluteRow(idx)
        session.findById(f"wnd[0]/usr/tblSAPL0ORGCORETCTRL_V_TVTW/ctxtV_TVTW-VTWEG[0,{idx}]").text = distribution_channel
        session.findById(f"wnd[0]/usr/tblSAPL0ORGCORETCTRL_V_TVTW/txtV_TVTW-VTEXT[1,{idx}]").text = Description

    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()


def main():
     
     session = get_session()
     goto_tcode(session) 
     distribution_channel(session) 

if __name__ == "__main__":
         main()

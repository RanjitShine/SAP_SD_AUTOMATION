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
   
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nv_tspa"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[1]/btn[5]").press()
    
 
def Division(session):
 
    df = pd.read_csv("test/Division.csv")
    df = df.reset_index(drop=True)
   
    for index, row in df.iterrows():
        Division = str(row["Division"]).strip()
        description = str(row["Description"]).strip()
       
        session.findById("wnd[0]/usr/tblSAPL0ORGCORETCTRL_V_TSPA").getAbsoluteRow(index)
        session.findById(f"wnd[0]/usr/tblSAPL0ORGCORETCTRL_V_TSPA/ctxtV_TSPA-SPART[0,{index}]").text = Division
        session.findById(f"wnd[0]/usr/tblSAPL0ORGCORETCTRL_V_TSPA/txtV_TSPA-VTEXT[1,{index}]").text = description
    
 
        time.sleep(1)
        print(f"Row {index+1} saved successfully")

    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()


def main():
     
     session = get_session()
     goto_tcode(session) 
     Division(session) 

if __name__ == "__main__":
         main()

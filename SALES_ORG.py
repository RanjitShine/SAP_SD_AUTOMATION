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
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nOVX5"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[1]/btn[5]").press()

def sales_orgs(session):
    csv_path = resource_path("test", "Sales_Org.csv")
    print("CSV path:", csv_path)

    df = pd.read_csv(csv_path, encoding="utf-8-sig")
    df = df.reset_index(drop=True)

    for idx, row in df.iterrows():
        vkorg         = str(row["Sales organization"]).strip()
        vkorg_name    = str(row["Sales organization Name"]).strip()
        currency      = str(row["statistics currency"]).strip()

        title         = str(row["Title"]).strip()
        addr_name     = str(row["Name"]).strip()
        search_term_1 = str(row["Search term 1"]).strip()
        search_term_2 = str(row["Search term 2"]).strip()
        street2       = str(row["Street 2"]).strip()
        street        = str(row["Street"]).strip()
        house_number  = str(row["House number"]).strip()
        postal_code   = str(row["Postal Code"]).strip()
        city          = str(row["City"]).strip()
        country       = str(row["Country"]).strip()
        region        = str(row["Region"]).strip()
        language      = str(row["Language"]).strip()
        telephone     = str(row["Telephone"]).strip()
        fax           = str(row["Fax"]).strip()
        email         = str(row["E-Mail"]).strip()
        Description   = str(row["TR-Description"]).strip()


        region = str(region).zfill(2)

        session.findById("wnd[0]/usr/ctxtV_TVKO-VKORG").text = vkorg
        session.findById("wnd[0]/usr/txtV_TVKO-VTEXT").text = vkorg_name
        session.findById("wnd[0]/usr/ctxtV_TVKO-WAERS").text = currency
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/cmbSZA1_D0100-TITLE_MEDI").key = "Company"
        session.findById("wnd[1]/usr/txtADDR1_DATA-NAME1").text      = addr_name
        session.findById("wnd[1]/usr/txtADDR1_DATA-STREET").text     = street
        # session.findById("wnd[1]/usr/txtADDR1_DATA-HOUSE_NUM1").text = house_number
        session.findById("wnd[1]/usr/txtADDR1_DATA-POST_CODE1").text = postal_code
        session.findById("wnd[1]/usr/txtADDR1_DATA-CITY1").text      = city
        session.findById("wnd[1]/usr/ctxtADDR1_DATA-COUNTRY").text   = country
        session.findById("wnd[1]/usr/ctxtADDR1_DATA-REGION").text    = region
        session.findById("wnd[1]/usr/txtADDR1_DATA-PO_BOX").text     = ""
        session.findById("wnd[1]/usr/txtADDR1_DATA-POST_CODE2").text = ""
        session.findById("wnd[1]/usr/txtADDR1_DATA-POST_CODE3").text = ""
        session.findById("wnd[1]/usr/txtSZA1_D0100-TEL_NUMBER").text = telephone
        session.findById("wnd[1]/usr/txtSZA1_D0100-MOB_NUMBER").text = ""
        session.findById("wnd[1]/usr/txtSZA1_D0100-FAX_NUMBER").text = fax
        session.findById("wnd[1]/usr/txtSZA1_D0100-SMTP_ADDR").text  = email
        session.findById("wnd[1]/usr/txtSZA1_D0100-SMTP_ADDR").setFocus()
        session.findById("wnd[1]/usr/txtSZA1_D0100-SMTP_ADDR").caretPosition = len(email)
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[2]/usr/txtKO013-AS4TEXT").text = Description
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        






def main():
     open_and_login() 
     session = get_session()
     goto_tcode(session) 
     sales_orgs(session) 
    

if __name__ == "__main__":
         main()

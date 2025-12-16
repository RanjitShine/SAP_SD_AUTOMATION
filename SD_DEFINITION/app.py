import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import SALES_ORG
import DISTRIBUTION_CHANNEL
import Division as DivisionModule
import SALES_OFFICE
import SALES_GROUP
import SHIPPING_POINT

FOLDER_NAME = "data"   # folder where CSVs will be saved

def automate_all():

    run_sales_org()
    run_distribution_channel()
    run_division()
    run_sales_office()
    run_sales_group()
    run_shipping_point()

def get_base_dir():

    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def convert_excel_to_csv_all_sheets(excel_path):

    base_dir = get_base_dir()
    output_folder = os.path.join(base_dir, FOLDER_NAME)
    os.makedirs(output_folder, exist_ok=True)
    xls = pd.ExcelFile(excel_path)

    for sheet in xls.sheet_names:

        df = pd.read_excel(excel_path, sheet_name=sheet, header=2)
        df = df[df.iloc[:, 0] != "Length"]
        safe_name = "".join(c if c.isalnum() else "_" for c in sheet)
        csv_path = os.path.join(output_folder, f"{safe_name}.csv")

        # Save CSV
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")

    return output_folder


def select_and_convert():
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if not file_path:
        return

    try:
        out_folder = convert_excel_to_csv_all_sheets(file_path)
        messagebox.showinfo("Done", f"All sheets converted.\nSaved in:\n{out_folder}")
    except Exception as e:
        messagebox.showerror("Error", f"Conversion failed:\n{e}")


# ------------ SAP Script wrappers (NO subprocess) ------------

def run_sales_org():
    try:
        SALES_ORG.main()   # <-- no base_dir argument
        # messagebox.showinfo("Done", "Sales Org automation completed.")
    except Exception as e:
        messagebox.showerror("Error", f"Sales Org failed:\n{e}")


def run_distribution_channel():
    try:
        DISTRIBUTION_CHANNEL.main()   # <-- no argument
        # messagebox.showinfo("Done", "Distribution Channel automation completed.")
    except Exception as e:
        messagebox.showerror("Error", f"Distribution Channel failed:\n{e}")


def run_division():
    try:
        DivisionModule.main()   # <-- no argument
        # messagebox.showinfo("Done", "Division automation completed.")
    except Exception as e:
        messagebox.showerror("Error", f"Division failed:\n{e}")


def run_sales_office():
    try:
        SALES_OFFICE.main()   # <-- no argument
        # messagebox.showinfo("Done", "Sales Office automation completed.")
    except Exception as e:
        messagebox.showerror("Error", f"Sales Office failed:\n{e}")


def run_sales_group():
    try:
        SALES_GROUP.main()   # <-- no argument
        # messagebox.showinfo("Done", "Sales Group automation completed.")
    except Exception as e:
        messagebox.showerror("Error", f"Sales Group failed:\n{e}")

def run_shipping_point():
    try:
        SHIPPING_POINT.main()   # <-- no argument
        # messagebox.showinfo("Done", "Shipping point automation completed.")
    except Exception as e:
        messagebox.showerror("Error", f"Shipping point failed:\n{e}")


def main():
    root = tk.Tk()
    root.title("SAP Automation")

    tk.Button(root,text="1) Select Excel & Convert to CSV (folder: test)",command=select_and_convert,width=45).pack(padx=30, pady=15)
    tk.Button(root, text="Automate ALL", command=automate_all).pack(padx=30, pady=5)
    tk.Button(root, text="Run Sales Org Script", command=run_sales_org).pack(padx=30, pady=5)
    tk.Button(root, text="Run Distribution Channel", command=run_distribution_channel).pack(padx=30, pady=5)
    tk.Button(root, text="Run Division", command=run_division).pack(padx=30, pady=5)
    tk.Button(root, text="Run Sales Office", command=run_sales_office).pack(padx=30, pady=5)
    tk.Button(root, text="Run Sales Group", command=run_sales_group).pack(padx=30, pady=5)
    tk.Button(root, text="Run Shipping Point", command=run_shipping_point).pack(padx=30, pady=5)


    root.mainloop()

if __name__ == "__main__":
    main()

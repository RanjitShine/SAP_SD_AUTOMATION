import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import ASSIGN_SALES_ORG_DISTRIBUTION



    # run_sales_org()
    # run_distribution_channel()
    # run_division()
    # run_sales_office()
    # run_sales_group()
    # run_shipping_point()

def assign_1():
    try:
        ASSIGN_SALES_ORG_DISTRIBUTION.main()   # <-- no argument
        messagebox.showinfo("Done", "Shipping point automation completed.")
    except Exception as e:
        messagebox.showerror("Error", f"Shipping point failed:\n{e}")


def main():
    root = tk.Tk()
    root.title("SAP Automation")

    tk.Button(root,text="Assign sales org and Distribution channel",command=assign_1,width=45).pack(padx=30, pady=15)
    root.mainloop()

if __name__ == "__main__":
    main()
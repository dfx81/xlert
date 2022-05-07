# XLert by dfx
# TODO: Refactor code

from datetime import datetime
from tkinter import *
from tkinter import filedialog
from tkinter.ttk import *
from tkinter import messagebox
import pandas as pd
import sys
import os
import os.path
import winshell

def main(path, columns, reminder_in_days=15):
    dues = 0
    data = pd.read_excel(path, usecols=columns)
    today = datetime.today()
    
    for item in data.to_records():
        try:
            due_date = datetime.strptime(str(item[1]), "%d.%m.%y")
            
            if (due_date - today).days <= reminder_in_days:
                dues += 1

        except ValueError as err:
            continue

    message = f"You have {dues} alerts."

    if (dues == 0):
        message += "\nNo further action required."
        messagebox.showinfo(title="XLert by dfx", message=message)
    else:
        message += "\nFurther action is required."
        messagebox.showwarning(title="XLert by dfx", message=message)
        if messagebox.askyesno(title="XLert by dfx", message="Do you want to open the excel file?"):
            os.startfile(path)

def open_fpicker(label):
    filetypes = (
        ("Excel files", ["*.xls", "*.xlsm", "*.xlsx"]),
    )

    filepath = filedialog.askopenfilename(filetypes=filetypes)

    if filepath:
        label["text"] = filepath

def show_credit():
    credit_text  = "XLert - Excel Alert v1.0\n\n"
    credit_text += "Programmer/Designer - Danial Fitri Ghazali (dfx)\n"
    credit_text += "Project Manager - Haikal Ghazali\n\n"
    credit_text += "For more info, visit https://github.com/dfx81/xlert"

    messagebox.showinfo(master=window, title="Credits", message=credit_text)

def create_alert(file_txt, col_inp, days_inp):
    err = 0
    err_message = ""

    if not os.path.exists(file_txt["text"]):
        err += 1
        err_message += "- Please select an excel file\n"
    if col_inp.get().isdigit() or not col_inp.get().isalpha():
        err += 1
        err_message += "- Please enter a single column letter\n"
    if not days_inp.get().isdigit():
        err += 1
        err_message += "- Please enter a valid number of days\n"

    if err:
        messagebox.showwarning(title="Error", message=err_message)
    else:
        file = file_txt["text"]
        col = col_inp.get().upper()
        days = int(days_inp.get())

        exepath = ""

        if getattr(sys, 'frozen', False):
            exepath = os.path.abspath(sys.executable)
        else:
            exepath = os.path.abspath(__file__)

        link_path = os.path.join(winshell.startup(), f"{file.split('/')[-1]}{col}{days}.lnk")

        with winshell.shortcut(link_path) as lnk:
            lnk.path = exepath
            lnk.arguments = f'"{file}" "{col}" "{days}"'

        raise SystemExit(0)

if __name__ == "__main__":
    if len(sys.argv) < 4:
        window = Tk()
        window.title("XLert - Excel Alert")
        window.geometry(f"320x140+{window.winfo_screenwidth() // 2 - (320 // 2)}+{window.winfo_screenheight() // 2 - (156 // 2)}")
        window.resizable(False, False)

        frm_label = Frame(master=window)
        lbl_instr = Label(master=frm_label, text="Select an excel file:")
        lbl_instr.pack(side=LEFT)
        frm_label.pack(fill=BOTH, padx=8, pady=(8, 0))

        frm_input = Frame(master=window)
        lbl_path = Label(master=frm_input, text="No File Selected", width=32, foreground="grey")
        lbl_path.pack(side=LEFT, padx=(0, 4))
        btn_file = Button(master=frm_input, text="Browse", command=lambda: open_fpicker(lbl_path))
        btn_file.pack(fill=Y, side=RIGHT)
        frm_input.pack(fill=BOTH, padx=8, pady=(0, 4))

        frm_input = Frame(master=window)
        lbl_instr = Label(master=frm_input, text="Date column (dd.mm.yy):")
        lbl_instr.pack(side=LEFT, padx=(0, 4))
        ent_col = Entry(master=frm_input)
        ent_col.pack(fill=X, side=RIGHT)
        frm_input.pack(fill=BOTH, padx=8, pady=(0, 4))

        frm_input = Frame(master=window)
        lbl_instr = Label(master=frm_input, text="Days before due date (days):")
        lbl_instr.pack(side=LEFT, padx=(0, 4))
        ent_days = Entry(master=frm_input)
        ent_days.pack(fill=X, side=RIGHT)
        frm_input.pack(fill=BOTH, padx=8, pady=(0, 4))

        frm_input = Frame(master=window)
        btn_submit = Button(master=window, text="Create Alert", command=lambda: create_alert(lbl_path, ent_col, ent_days))
        btn_submit.pack(side=RIGHT, padx=(0, 8), pady=(0, 8))
        btn_cred = Button(master=window, text="Credits", command=show_credit)
        btn_cred.pack(side=RIGHT, padx=(8, 4), pady=(0, 8))
        frm_input.pack(fill=BOTH, padx=8, pady=8)

        window.mainloop()
        
        raise SystemExit(0)
    
    path = sys.argv[1]
    columns = sys.argv[2]
    reminder_in_days = int(sys.argv[3])
    
    main(path, columns, reminder_in_days)
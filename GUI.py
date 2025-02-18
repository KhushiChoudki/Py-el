import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from datetime import date
import openpyxl

def insert_row():
    name = name_entry.get()
    dob = (cal.get_date())
    id = int(id_entry.get())
    des = des_entry.get()
    # Insert row into excel
    path = r"C:\Users\Khushi\Desktop\Py el\Employee.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [name, dob, id, des]
    sheet.append(row_values)
    workbook.save(path)
    # Insert treeview
    treeview.insert('', tk.END, values=row_values)
    # Clear
    name_entry.delete(0, "end")
    name_entry.insert(0, "Employee Name")
    cal.delete(0, "end")
    cal.insert(0, "Employee DOB: ")
    id_entry.delete(0, "end")
    id_entry.insert(0, "Employee ID")
    des_entry.delete(0, "end")
    des_entry.insert(0, "Employee Designation")

def load_data():
    path = r"C:\Users\Khushi\Desktop\Py el\Employee.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    list_values = list(sheet.values)
    print(list_values)
    for col_name in list_values[0]:
        treeview.heading(col_name, text=col_name, anchor="center")
    for value_tuple in list_values[1:]:
        treeview.insert('', tk.END, values=value_tuple)

def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")

root = tk.Tk()
style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")
frame = ttk.Frame(root)
frame.grid(row=0, column=0, sticky='nsew')
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)
widgets_frame = ttk.LabelFrame(frame, text="Insert Row")
widgets_frame.grid(row=0, column=0, padx=20, pady=10, sticky='nsew')
widgets_frame.grid_columnconfigure(0, weight=1)
name_entry = ttk.Entry(widgets_frame)
name_entry.insert(0, "Employee Name")
name_entry.bind("<FocusIn>", lambda e: name_entry.delete('0', 'end'))
name_entry.grid(row=0, column=0, padx=5, pady=[0, 5], sticky="ew")
cal = DateEntry(widgets_frame, selectmode='day')
cal.insert(0, "Employee DOB: ")
cal.bind("<FocusIn>", lambda e: cal.delete('0', 'end'))
cal.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
id_entry = ttk.Entry(widgets_frame)
id_entry.insert(0, "Employee ID")
id_entry.bind("<FocusIn>", lambda e: id_entry.delete('0', 'end'))
id_entry.grid(row=2, column=0, padx=5, pady=5, sticky="ew")
des_entry = ttk.Entry(widgets_frame)
des_entry.insert(0, "Employee Designation")
des_entry.bind("<FocusIn>", lambda e: des_entry.delete('0', 'end'))
des_entry.grid(row=3, column=0, padx=5, pady=5, sticky="ew")
button = ttk.Button(widgets_frame, text="Insert", command=insert_row)
button.grid(row=4, column=0, padx=5, pady=5, sticky="ew")
seperator = ttk.Separator(widgets_frame)
seperator.grid(row=5, column=0, padx=[20, 10], pady=10, sticky="ew")
mode_switch = ttk.Checkbutton(widgets_frame, text="Mode", style="Switch", command=toggle_mode)
mode_switch.grid(row=6, column=0, padx=5, pady=10, sticky="nsew")
treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10, sticky='nsew')
frame.grid_rowconfigure(0, weight=1)
frame.grid_columnconfigure(1, weight=1)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")
cols = ["Name", "DOB/TIME", "ID", "Designation"]
treeview = ttk.Treeview(treeFrame, show="headings", yscrollcommand=treeScroll.set, columns=cols, height=13)
treeview.column("Name", anchor="center", width=100)
treeview.column("DOB/TIME", anchor="center", width=100)
treeview.column("ID", anchor="center", width=50)
treeview.column("Designation", anchor="center", width=100)
treeview.pack(fill='both', expand=True)
treeScroll.config(command=treeview.yview)
load_data()
root.mainloop()
import tkinter as tk
from auto_report import main
from tkinter import ttk
from TkinterDnD2 import DND_FILES, TkinterDnD

def drop_inside_box(event):
    batch_path.set(event.data)

def sort_report():
    # print(batch_path.get())
    #tkinter adds curly braces {} to batch_path.get() return value due to tcl list encoding
    #.strip('{}') is a temporary workaround until finding where batch_path is set as a list
    main(batch_path.get().strip('{}'))

window = TkinterDnD.Tk()
window.title('Auto Report')
window.geometry('800x500')
window.resizable(False, False)

#PY_VAR (requires batch_path.set() and batch_path.get())
batch_path = tk.StringVar(master = window)

file_entry = ttk.Entry(master = window, textvariable = batch_path)
file_entry.pack()
file_entry.drop_target_register(DND_FILES)
file_entry.dnd_bind("<<Drop>>", drop_inside_box)

sort_button = ttk.Button(master = window, text = 'Sort Batch Report', command = sort_report)
sort_button.pack()

# list_button = ttk.Button(master = window, text = 'Coffee List')
# list_button.pack()

window.mainloop()

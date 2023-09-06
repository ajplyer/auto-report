import tkinter as tk
from auto_report import main
from tkinter import ttk
from TkinterDnD2 import DND_FILES, TkinterDnD
from datetime import datetime
from os import path, startfile
        
#Entry event sets batch_path on drop
def drop_inside_box(event):
    batch_path.set(event.data)
    sort_button.config(state = 'normal')

#Sort button command runs auto_report.py using batch_path
def sort_report():
    global sorted_file
    #auto_report saves sorted file to same directory as original batch report, sorted file is named using current month, day, year, hour, minute
    sorted_file = path.dirname(batch_path.get().strip('{}')) + '/batch ' + datetime.now().strftime('%m%d%y%H%M') + '.xlsx'

    #Tkinter adds curly braces {} to batch_path.get() return value due to tcl list encoding
    #.strip('{}') is a temporary workaround until finding where batch_path is set as a list
    main(batch_path.get().strip('{}'))

    open_file_button.config(state = 'normal')

def open_sorted():
    startfile(sorted_file)

def clear_entry():
    file_entry.delete(0, tk.END)
    sort_button.config(state = 'disabled')
    open_file_button.config(state = 'disabled')

#Root window 
window = TkinterDnD.Tk()
window.title('Auto Report')
window.geometry('400x500')
window.resizable(False, False)

batch_path = tk.StringVar(window)

file_entry = ttk.Entry(window, textvariable = batch_path, width = 50)
file_entry.pack(ipady = 100, pady = 25)
file_entry.drop_target_register(DND_FILES)
file_entry.dnd_bind("<<Drop>>", drop_inside_box)

clear_entry_button = ttk.Button(window, text = 'Clear Entry', command = clear_entry)
clear_entry_button.pack(pady = 5, ipady = 5)

#Sort button begins disabled until a batch report is inserted into entry box
sort_button = ttk.Button(window, text = 'Sort Batch Report', width = 25, command = sort_report, state = 'disabled')
sort_button.pack(pady = 25, ipady = 5)

#Open file button begins disabled until auto_report.py is finished running
open_file_button = ttk.Button(window, text = 'Open Sorted Batch', width = 25, command = open_sorted, state = 'disabled')
open_file_button.pack(pady = 5, ipady = 5)

window.mainloop()

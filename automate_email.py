#  automate_email.py
try:
    # Python 2
    import Tkinter as tk
    import ttk
    from Tkinter import filedialog
    from tkinter import StringVar, Entry
    from tkFileDialog import askopenfilename
except ImportError:
    # Python 3
    import tkinter as tk
    from tkinter import filedialog
    from tkinter import ttk
    from tkinter import StringVar, Entry
    from tkinter.filedialog import askopenfilename

import pandas as pd
import numpy as np
from pandas import ExcelFile
from pandas import ExcelWriter
from email_class import Email

pd.options.display.max_rows = 999

# --- classes ---

class Orders:
    def __init__(self, order_number, email, name):
        self.order_number = order_number
        self.email = email
        self.name = name

class MyWindow:

    def __init__(self, parent):
        self.parent = parent

        self.filename = None
        self.df = None
        self.folder_selected = None

        self.parent.grid_rowconfigure(0, weight=1)
        self.parent.grid_columnconfigure(0, weight=1)

        self.text_email = tk.Text(self.parent, width=100, height=30)
        self.text_email.grid(row=0, column=1, columnspan = 2) 

        self.text_sent = tk.Text(self.parent, width=100, height=30)
        self.text_sent.grid(row=1, column=1, columnspan = 2) 

        self.email = StringVar() #Email variable
        self.emailEntry = Entry(self.parent, textvariable=self.email).grid(row=2, column=1, pady=(10,0)) 
        self.submit = ttk.Button(self.parent, text='Enter Email').grid(row=3, column=1) 

        self.password = StringVar() #Password variable
        self.passEntry = Entry(self.parent, textvariable=self.password, show='*').grid(row=4, column=1, pady=10) 
        self.submit = ttk.Button(self.parent, text='Enter Password').grid(row=5, column=1, pady=(0,20))  

        self.loademail_button = ttk.Button(self.parent, text='LOAD EMAIL LIST', command=self.load)
        self.loademail_button.grid(row=2, column=2, pady=(20,10)) 

        self.loadimages_button = ttk.Button(self.parent, text='LOAD IMAGES', command=self.display)
        self.loadimages_button.grid(row=3, column=2) 

        self.sendemail_button = ttk.Button(self.parent, text='SEND EMAILS', command=self.send)
        self.sendemail_button.grid(row=4, column=2) 


    
    #function for load button
    def load(self):

        name = askopenfilename(filetypes=[('Excel', ('*.xls', '*.xlsx')), ('CSV', '*.csv',)])

        if name:
            if name.endswith('.csv'):
                self.df = pd.read_csv(name)
            else:
                self.df = pd.read_excel(name)

            self.filename = name
            self.text_email.insert('end', self.filename + '\n')
            self.text_email.insert('end', str(self.df) + '\n')


            
    #function for display button
    def display(self):
        # ask for file if not loaded yet
        if self.df is None:
            self.load()

        # display if loaded
        if self.df is not None:
            self.folder_selected = filedialog.askdirectory()


    #function for save button
    def send(self):
        if self.df is None:
            self.display()

        if self.folder_selected is None:
            self.display()

        orders = [(Orders(row.Order,row.Email,row.Name)) for index, row in self.df.iterrows() ]
                
        for x in orders:
            try:
                Email(self.email.get(), self.password.get()).send_email(x.name, x.email, x.order_number, self.folder_selected)

            except (AttributeError, TypeError):
                    self.text_sent.insert('end', "Try Again" + '\n') 

            self.text_sent.insert('end', "Email Order " + str(x.order_number) + " sent to " + str(x.email) + " " + str(x.name) + '\n')




# --- main ---

if __name__ == '__main__':
    root = tk.Tk()
    top = MyWindow(root)
    root.mainloop()


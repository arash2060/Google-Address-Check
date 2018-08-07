print('Loading ...')
from googlemaps import Client
from numpy import number, issubdtype
from os import path, remove
from configparser import RawConfigParser
from tkinter import *
from tkinter import messagebox
import random
waitlist = ['Wanna go get some press covfefe... and come back? Just kidding!',
            'This will be ready soon! No need for another Snapchat...',
            'Good things happen after you wait a gazillion years, OR you make them happen!',
            'A few more seconds... ...']
print(random.choice(waitlist))
from tkinter.ttk import *
from pandas import ExcelFile, ExcelWriter, merge
from datetime import datetime
import time
from logging import basicConfig, DEBUG

print('Almost ready!')
# Logging includes private information to the log i.e. Google API Key and
# the addresses searched. You should add ./Address_Check.log to .gitignore
basicConfig(filename='./Address_Check.log', level=DEBUG)


from tkinter import ttk

import threading
import queue
import time

fields = ['Google API Key', 'Input File']
combos = ['Business Name:', 'Legal Name:', 'Street Address:', 'Street Number:',
          'Street Name:', 'City/Borough:', 'Zipcode:', 'Boro Code:', 'State:']


class App(Tk):

    # def __init__(self):
    #     tk.Tk.__init__(self)
    #     self.queue = queue.Queue()
    #     self.listbox = tk.Listbox(self, width=20, height=5)
    #     self.progressbar = ttk.Progressbar(self, orient='horizontal',
    #                                        length=300, mode='determinate')
    #     self.button = tk.Button(self, text="Start", command=self.spawnthread)
    #     self.listbox.pack(padx=10, pady=10)
    #     self.progressbar.pack(padx=10, pady=10)
    #     self.button.pack(padx=10, pady=10)

    def __init__(self, fields, combos):
        Tk.__init__(self)
        # Quick workaround: Window background is white on Mac while buttons ... are grey.
        #self.configure(background='grey91')

        self.title("Google Geocoding")
        self.ents, self.row = self.makeform(fields)
        self.bind('<Return>', (lambda event, e=self.ents: fetch(e)))
        self.b1 = Button(self.row, text='Browse...', command=self.browsexlsx)
        self.b1.pack(side=LEFT, padx=5, pady=5)
        self.b2 = Button(self.row, text='Load Sheets', command=self.loadxlsx)
        self.b2.pack(side=LEFT, padx=5, pady=5, anchor=W)
        self.row.pack(side=TOP, fill=X, padx=5, pady=5)

        self.row = Frame(self)
        self.lbl_sheet = Label(self.row, text="Choose Input Sheet:", width=20, anchor='w')
        self.sheet_combo = Combobox(self.row, width=20, state='disabled')
        self.lbl_frow = Label(self.row, text="First Row?", width=10, anchor='w')
        self.frow = Entry(self.row, width=5, state='disabled')
        self.lbl_sheet.pack(side=LEFT, padx=5, pady=5)
        self.sheet_combo.pack(side=LEFT, expand=YES, fill=X)
        self.lbl_frow.pack(side=LEFT, padx=5, pady=5)
        self.frow.pack(side=LEFT, expand=YES, fill=X)

        self.b3 = Button(self.row, text='Load Fields', command=self.loadfields, state='disabled')

        self.b3.pack(side=LEFT, padx=5, pady=5, anchor='w')
        self.row.pack(side=TOP, fill=X, padx=5, pady=5)

        self.combs, self.row2 = self.makecomboboxes(combos)

        self.row = Frame(self)
        self.second_run_state = BooleanVar()
        self.second_run_state.set(True) #set check state
        self.chk = Checkbutton(self.row, text="Run twice using Address and Trade & Address", var=self.second_run_state, state='disabled')
        self.chk.pack(side=LEFT, padx=15, pady=5, anchor='w')
        self.row.pack(side=TOP, fill=X, padx=5, pady=5)

        self.row = Frame(self)
        self.b4 = Button(self.row, text='Geocode', command=lambda: Geocode(DFrame, combs), state='disabled')
        self.out_lbl = Label(self.row, text="Output File", width=20, anchor='w')
        self.output = Entry(self.row, width=40, state='disabled')
        self.out_lbl.pack(side=LEFT, padx=5, pady=5)
        self.output.pack(side=LEFT, expand=YES, fill=X)
        self.b4.pack(side=LEFT, padx=5, pady=5, anchor='w')
        self.row.pack(side=TOP, fill=X, padx=5, pady=5)

        self.row = Frame(self)
        self.status = StringVar()
        self.status.set("Status: waiting for user input ...")
        self.status_bar = Label(self.row, textvariable=self.status, bo=0.1,
                           relief=SUNKEN, anchor='w')
        self.status_bar.pack(side=BOTTOM, fill=X, padx=5, pady=5)
        self.row.pack(side=TOP, fill=X, padx=5, pady=5)


    def fetch(self, entries):
        for entry in entries:
            field = entry[0]
            text = entry[1].get()
            print('%s: "%s"' % (field, text))

    def makeform(self, fields):
        self.entries = []
        for field in fields:
            self.row = Frame(self)
            self.lab = Label(self.row, width=20, text=field, anchor='w')
            self.ent = Entry(self.row, width=60)
            if field != fields[-1]:
                self.row.pack(side=TOP, fill=X, padx=5, pady=5)
            self.lab.pack(side=LEFT, padx=5, pady=5)
            self.ent.pack(side=LEFT, expand=YES, fill=X)
            self.entries.append((field, self.ent))
        return self.entries, self.row

    def makecomboboxes(self, combos):
        comboboxes = []
        for combo in combos:
            row = Frame(self)
            lab = Label(row, width=15, text=combo, anchor='w')
            ent = Combobox(row, width=20, state='disabled')
            row.pack(side=TOP, fill=X, padx=5, pady=5)
            lab.pack(side=LEFT, padx=15, pady=5)
            ent.pack(side=LEFT, padx=10, expand=YES, fill=X)
            comboboxes.append((combo, ent))
        return comboboxes, row

    def set_text(self, text):
        self.ents[1][1].delete(0, END)
        self.ents[1][1].insert(0, text)
        return

    def keyfunction(self, x):
        '''
            used to sort mixed list of column names
        '''
        v = x
        if isinstance(v, int):
            v = '0%d' % v
        return v

    def choose_default(self, i, collist, field):
        '''
            If default fields exist on the excel sheet,
            automatically choose them in the combobox values
        '''
        if field in ['streetnumber', 'streetname'] and 'originaladdress' in collist:
            self.combs[i][1].current(0)
        elif field in sorted(collist, key=self.keyfunction):
            location = [i for i,x in enumerate(sorted(collist, key=self.keyfunction)) if x == field][0]
            self.combs[i][1].current(location)
        else:
            self.combs[i][1].current(0)

    def browsexlsx(self):
        from tkinter.filedialog import askopenfilename
        from os import path

        
        # self.withdraw()
        # possible other option: multifile=1
        # '.xls*' doesn't work on Mac.
        filenames = askopenfilename(parent=self, filetypes=[('Excel files', ['.xlsx','.xls']), ('All files', '.*')],
                                    initialdir=path.dirname(r"Z:\EAD\DOL Data\QCEW to RPAD address merge\forgbat"))
        # response = self.tk.splitlist(filenames)
        # for f in response:
        #     print(f)
        print(filenames)
        self.set_text(filenames)


    def loadxlsx(self):
        '''
        Get the sheet names in the chosen excel file.
        '''
        filename = self.ents[1][1].get()
        if (filename.endswith(".xlsx") or filename.endswith(".xls")):
            # Read in the Load The Sheets.
            print("Processing: %s" % filename.encode().decode())
            f = path.basename(filename)
            self.status.set("Status: loading sheets of %s" % f.encode().decode())
            adds = ExcelFile(filename)
            print("This Excel file includes these sheets: %s" % adds.sheet_names)
            self.sheet_combo['state'] = 'enabled'
            self.frow['state'] = 'enabled'
            self.b3['state'] = 'enabled'
            self.output['state'] = 'enabled'
            self.sheet_combo['values'] = adds.sheet_names
            self.output.delete(0, END)
            self.output.insert(0, filename.replace(".xls", "_out.xls"))
            self.sheet_combo.current(0)
            self.status.set("Status: Choose which sheet of %s to process and press 'Load Fields'" % f.encode().decode())
        else:
            messagebox.showinfo(title="Not Excel File", message="Enter and Excel File")


    def loadfields(self):
        '''
        Get the variable names in the chosen excel sheet
        '''

        filename = self.ents[1][1].get()
        f = path.basename(filename)
        self.status.set("Status: loading data and column names of %s" % f.encode().decode())
        self.adds = ExcelFile(filename)
        self.sheet = self.sheet_combo.get()
    #   if first row is not entered, assume 1 and set the form to 1.
        if self.frow.get() == "":
            self.frow.insert(0, 1)
            self.first_row = 1
        else:
            self.first_row = int(self.frow.get())
        print("%s and %s onwards chosen." % (self.sheet, self.first_row))
        df = self.adds.parse(self.sheet, skiprows=self.first_row - 1)
        #print(df.columns.values)
        print("There are %s observations on this file." % len(df.index))
        ['Business Name:', 'Street Number:',
         'Street Name:', 'City/Borough:', 'Zipcode:', 'Boro Code:']
        self.defaults = {0: 'trade',
                    1: 'legal',
                    2: 'originaladdress',
                    3: 'streetnumber',
                    4: 'streetname',
                    5: 'Borough',
                    6: 'pzip',
                    7: 'boro',
                    8: 'state'}
        for i in range(len(combos)):
            self.collist = list(df.columns.values)
            self.collist.append("")
            self.combs[i][1]['state'] = 'enabled'
            self.combs[i][1]['values'] = sorted(self.collist, key=self.keyfunction)
            self.choose_default(i, self.collist, self.defaults[i])
        self.chk['state'] = 'enabled'
        self.b4['state'] = 'enabled'
    #    print(combs[0][0], df[combs[0][1].get()].head(10))
        self.status.set("Status: Choose address fields, optionally edit output file, and press 'Geocode'")
        self.DFrame = df
        return df




    # def spawnthread(self):
    #     self.button.config(state="disabled")
    #     self.thread = ThreadedClient(self.queue)
    #     self.thread.start()
    #     self.periodiccall()

    # def periodiccall(self):
    #     self.checkqueue()
    #     if self.thread.is_alive():
    #         self.after(100, self.periodiccall)
    #     else:
    #         self.button.config(state="active")

    # def checkqueue(self):
    #     while self.queue.qsize():
    #         try:
    #             msg = self.queue.get(0)
    #             self.listbox.insert('end', msg)
    #             self.progressbar.step(25)
    #         except Queue.Empty:
    #             pass


class ThreadedClient(threading.Thread):

    def __init__(self, queue):
        threading.Thread.__init__(self)
        self.queue = queue

    def run(self):
        for x in range(1, 5):
            time.sleep(2)
            msg = "Function %s finished..." % x
            self.queue.put(msg)


if __name__ == "__main__":
    app = App(fields, combos)
    app.mainloop()
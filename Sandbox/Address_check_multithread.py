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
        self.b4 = Button(self.row, text='Geocode', command=lambda: self.Geocode(self.DFrame, self.combs), state='disabled')
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

    def Geocode(self, df, combs):
        '''
            Description: Geocodes the records that were not found in
            GBAT using Google API.
            By default all the years are in the exclude list so that
            we don't redo any years and don't pay Google twice.
        '''
        self.status.set("Status: began geocoding")

        # this needs to be retreived just after the geocode button is pressed.
        output = self.output.get()

        self.google_api_key2 = self.ents[0][1].get()

        try:
            self.gmaps = Client(key=self.google_api_key2)
        except:
            try:
                config = RawConfigParser()
                config.read('./API_Keys.cfg')
                self.google_api_key = config.get('Google', 'QCEW_API_Key')
                self.gmaps = Client(key=google_api_key)
            except:
                messagebox.showinfo("Can't Connect to Google", "Oups! Please check \
that your API key is valid, doesn't have leading/trailing spaces, and you are \
connected to internet! \nYour API key looks like vNIXE0xscrmjlyV-12Nj_BvUPaw")
                return None

        # test
        geocode_result = self.gmaps.geocode("Indepenent Buget Office, \
New York, NY 10003")
        print(geocode_result[0]['geometry']['location']['lat'],
              geocode_result[0]['geometry']['location']['lng'])

        # print(geocode_result)
        # print("DFrame\n\n", df.head())
        # Dictionary of boros to be used in for Addresses.
        Boros = {1: 'Manhattan', 2: 'Bronx', 3: 'Brooklyn',
                 4: 'Queens', 5: 'Staten Island', 9: 'New York'}

        df.fillna('')
        # trade2 is trade name when available, and legal name when not.
        df['trade2'] = ""
        #   trade
        if combs[0][1].get() != "":
            df.loc[df[combs[0][1].get()].fillna('') != '', 'trade2'] = df[combs[0][1].get()]
            #   legal
            if combs[1][1].get() != "":
                df.loc[df[combs[0][1].get()].fillna('') == '', 'trade2'] = df[combs[1][1].get()]
            else:
                df.loc[df[combs[0][1].get()].fillna('') == '', 'trade2'] = ""
        else:
            df['trade2'] = ""

        # Handle missing fields.
        if combs[2][1].get() == '' and (combs[3][1].get() == '' or combs[4][1].get() == ''):
            messagebox.showinfo("No Address!", "Either 'Street Address' or ''Street Number' and 'Street Name'' are required.")
            return None
        elif combs[3][1].get() != '' or combs[4][1].get() != '':
            df['Generated_streetaddress'] = df[combs[3][1].get()] + df[combs[4][1].get()]
        else:
            df['Generated_streetaddress'] = df[combs[2][1].get()]

        if combs[7][1].get() == '':
            if combs[5][1].get() == '':
                df['no_boro'] = 1
                vals = combs[7][1]['values'] + ('no_boro',)
                combs[7][1]['values'] = sorted(vals, key=self.keyfunction)
                choose_default(7, vals, 'no_boro')
                df['no_city'] = 'New York'
            else:
                df['no_city'] = df[combs[5][1].get()]
        else:
            df['no_boro'] = df[combs[7][1].get()]
            df['no_city'] = df['no_boro'].map(Boros)

        if combs[6][1].get() == '':
            df['no_zip'] = ''
            vals = combs[7][1]['values'] + ('no_zip',)
            combs[6][1]['values'] = sorted(vals, key=self.keyfunction)
            self.choose_default(6, vals, 'no_zip')
        else:
            if issubdtype(df[combs[6][1].get()].dtype, number):
                df['no_zip'] = df[combs[6][1].get()].round(0)
            else:
                df['no_zip'] = df[combs[6][1].get()]

        if combs[8][1].get() == '':
            df['no_state'] = 'NY'
            vals = combs[7][1]['values'] + ('no_state',)
            combs[6][1]['values'] = sorted(vals, key=self.keyfunction)
            self.choose_default(6, vals, 'no_state')
        else:
            df['no_state'] = df[combs[8][1].get()]




        # trade(or legal) name + Original Address + City, State, Zip
        df['temp_add'] = ""
        #   originaladdress field
        df['Generated_streetaddress'] = df['Generated_streetaddress'].replace('*** NEED PHYSICAL ADDRESS ***', '')
        df.loc[df['Generated_streetaddress'].fillna('') != '',
               'temp_add'] = df['Generated_streetaddress'].fillna('') + ', '
        # the last bit maps boro code to borough name
        df['NameAddress'] = (df.trade2.fillna('') + ', ' +
                             df['temp_add'] + df['no_city'] +
                             ', ' + df['no_state'] + ' ' + df['no_zip'].apply(str))
        df['Address'] = (df['temp_add'] + df['no_city'] +
                         ', ' + df['no_state'] + ' ' + df['no_zip'].apply(str))
        df['NameAddress'].head()
        #   print(df.head())

        # drop some temp fields.
        df.drop(['temp_add', 'no_boro', 'no_zip', 'no_city', 'no_state'], axis=1, inplace=True)

        df = df.fillna('')

        # Run google API twice, for the Address only and Name+Address.
        df.reset_index(inplace=True)
        df['Goog_ID'] = df.index
        if (self.second_run_state.get() is True and
            combs[0][1].get() != "" and
            combs[1][1].get() != ""):
            add_list = ['Address', 'NameAddress']
            # Create a dataframe with unique observations on add_list
            df_unique = df[['Address', 'NameAddress', 'Generated_streetaddress', 'Goog_ID']].copy()
            df_unique.drop_duplicates(subset=add_list, keep="first", inplace=True)
            df_unique.reset_index(inplace=True)

            obs = (len(df_unique.index) * 2)
            print('There are %s unique observations to process...' % (obs))
        else:
            add_list = ['Address']
            # Create a dataframe with unique observations on add_list
            df_unique = df[['Address', 'Generated_streetaddress', 'Goog_ID']].copy()
            df_unique.drop_duplicates(subset=add_list, keep="first", inplace=True)
            df_unique.reset_index(inplace=True)

            obs = (len(df_unique.index) * 2)
            print('There are %s unique observations to process...' % (obs))
        df_unique['Gformatted_address0'] = ""
        df_unique['Glat0'] = 0
        df_unique['Glon0'] = 0
        df_unique['GPartial0'] = False
        df_unique['Gtypes0'] = ""
        df_unique['Gformatted_address1'] = ""
        df_unique['Glat1'] = 0
        df_unique['Glon1'] = 0
        df_unique['GPartial1'] = False
        df_unique['Gtypes1'] = ""
        df_unique['Borough0'] = ""
        df_unique['Borough1'] = ""
        i = -1
        # print(add_list)
        startTime = time.time()

        directory = path.dirname(output)
        self.count_query = 0
        for var in add_list:
            print('Started checking variable ', var)
            i += 1
            for index, row in df_unique.iterrows():
                # set index <= len(df_unique.index) to process all observations.
                if index <= 20 and not(var in ['Address'] and row['Generated_streetaddress'] == ''):
                    geocode_result = self.gmaps.geocode(row[var])
                    self.count_query += 1
                    self.status.set("Status: looking up observation %s of %s" % (self.count_query, obs))
                    if len(geocode_result) > 0:
                        if 'partial_match' in geocode_result:
                            df_unique.loc[df_unique.index == index,
                                   ['GPartial' + str(i),
                                    'Gformatted_address' + str(i),
                                    'Glat' + str(i),
                                    'Glon' + str(i), 'Gtypes' + str(i)
                                   ]
                                  ] = (
                                       geocode_result[0]['partial_match'],
                                       geocode_result[0]['formatted_address'],
                                       geocode_result[0]['geometry']['location']['lat'],
                                       geocode_result[0]['geometry']['location']['lng'],
                                       str(geocode_result[0]['types'])
                                      )
                        else:
                            df_unique.loc[df_unique.index == index,
                                    [
                                     'Gformatted_address' + str(i),
                                     'Glat' + str(i), 'Glon'+ str(i),
                                     'Gtypes'+ str(i)
                                     ]
                                    ] = (
                                         geocode_result[0]['formatted_address'],
                                         geocode_result[0]['geometry']['location']['lat'],
                                         geocode_result[0]['geometry']['location']['lng'],
                                         str(geocode_result[0]['types'])
                                        )
                    else:
                        df_unique.loc[df_unique.index == index,
                               [
                                'Gformatted_address' + str(i),
                                'Glat' + str(i), 'Glon' + str(i),
                                'Gtypes' + str(i)
                                ]
                               ] = ('Not Found', 0, 0, str([0]))
                    # Save a temporary recovery file
                    if self.count_query % 500 ==0:
                        writer = ExcelWriter(path.join(directory, "GOOGLE_recovery.xlsx").encode().decode())
                        df_unique.to_excel(writer, 'Sheet1')
                        writer.save()
                        print("Recovery File GOOGLE_recovery.xlsx Saved when index was %s at %s" % (self.count_query, datetime.now()))
            # Extract some address components. "THESE CAN BE IMPROVED"
            df_unique['Gzip' + str(i)] = 0
            pat2 = r".*NY ([0-9]{5}).*"
            repl0 = lambda m: m.group(1)
            df_unique['Gzip' + str(i)] = df_unique['Gformatted_address' + str(i)].str.replace(pat2, repl0)

            pat = r"([0-9\-]+)(.*), (.*?)(, NY [0-9]{5}.*)"
            repl1 = lambda m: m.group(1)
            repl2 = lambda m: m.group(2)
            repl3 = lambda m: m.group(3)
            df_unique['Gnumber' + str(i)] = df_unique['Gformatted_address' + str(i)].str.replace(pat, repl1)
            df_unique['Gstreet' + str(i)] = df_unique['Gformatted_address' + str(i)].str.replace(pat, repl2)
            df_unique['Borough' + str(i)] = df_unique['Gformatted_address' + str(i)].str.replace(pat, repl3)

            #print(df[['Gnumber'+ str(i),'Gformatted_address'+ str(i)]].head(10))
            #print(df.Borough1.head(10))

        # Save The Results
        df_unique['Both_Run_Same'] = (df_unique['Gformatted_address1'] ==
                                      df_unique['Gformatted_address0'])

        # Drop the temporary variables.
        df.drop(['Generated_streetaddress'], axis=1, inplace=True)
        df_unique.drop(['Address', 'Generated_streetaddress'], axis=1, inplace=True)
        # print(df_unique.columns.values)
        try:
            df_unique.drop(['NameAddress'], axis=1, inplace=True)
        except:
            None

        # Merge back unique addresses with geocodes with original df.
        result = merge(df, df_unique, on='Goog_ID', how='outer')
        result.drop(['index_y', 'Goog_ID'], axis=1, inplace=True)

        # ExcelFile(output)
        try:
            writer = ExcelWriter(output)
            result.to_excel(writer, (self.sheet_combo.get() + '_Geocoded')[0:-1])
            writer.save()
            message = output + "\n was successfully saved!\n There were %s queries made to Google Maps API" % (self.count_query)
            messagebox.showinfo('Success', message)
        except:
            writer = ExcelWriter(path.join(directory, "Google_Geocoded_" +
                                           time.strftime("%Y%m%d-%H%M%S") +
                                           ".xlsx").encode().decode())
            result.to_excel(writer, self.sheet_combo.get() + '_Geocoded')
            writer.save()
            message = ("Couldn't write to " + output + "\n saved Google_Geocoded_" +
                       time.strftime("%Y%m%d-%H%M%S") + ".xlsx to the same directory.\n There were %s queries made to Google Maps API" % (self.count_query))
            messagebox.showinfo('Access Denied!', message)
    #   remove the recovery file.
        remove(path.join(directory, "GOOGLE_recovery.xlsx").encode().decode())
        print('Processed data and saved: ', output)

        endTime = time.time()
        t = (endTime - startTime) / 60
        self.status.set('Status: Done. It took %s minutes to make %s queries.' % (round(t, 2), count_query))
        print('Took %s minutes to run.' % round(t, 2))



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
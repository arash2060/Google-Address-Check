'''
Possible Improvements:
    Removing duplicates should be done separately for addresses and address+names before geocoding. Currently, it's done
    all together.

    Bug: when a city is entered, no_boro doesn't get compiled. This causes an error without any error messages.
'''
print('Loading ...')
import logging
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
from pandas import ExcelFile, ExcelWriter, merge, to_numeric
import pandas as pd
from datetime import datetime
import time
import sys

class Logger(object):
    def __init__(self, filename="Default.log"):
        self.terminal = sys.stdout
        self.log = open(filename, "a")

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)

sys.stderr = Logger("errors.log")
sys.stdout = Logger("output.log")

logging.basicConfig(filename='./Address_Check.log', level=logging.DEBUG)


print('Almost ready!')
# Logging includes private information to the log i.e. Google API Key and
# the addresses searched. You should add ./Address_Check.log to .gitignore


fields = ['Google API Key', 'Input File']
combos = ['Business Name:', 'Legal Name:', 'Street Address:', 'Street Number:',
          'Street Name:', 'City/Borough:', 'Zipcode:', 'Boro Code:', 'State:']


def fetch(entries):
    for entry in entries:
        field = entry[0]
        text = entry[1].get()
        print('%s: "%s"' % (field, text))


def makeform(root, fields):
    entries = []
    for field in fields:
        row = Frame(root)
        lab = Label(row, width=20, text=field, anchor='w')
        ent = Entry(row, width=60)
        if field != fields[-1]:
            row.pack(side=TOP, fill=X, padx=5, pady=5)
        lab.pack(side=LEFT, padx=5, pady=5)
        ent.pack(side=LEFT, expand=YES, fill=X)
        entries.append((field, ent))
    return entries, row


def makecomboboxes(root, combos):
    comboboxes = []
    for combo in combos:
        row = Frame(root)
        lab = Label(row, width=15, text=combo, anchor='w')
        ent = Combobox(row, width=20, state='disabled')
        row.pack(side=TOP, fill=X, padx=5, pady=5)
        lab.pack(side=LEFT, padx=15, pady=5)
        ent.pack(side=LEFT, padx=10, expand=YES, fill=X)
        comboboxes.append((combo, ent))
    return comboboxes, row


def set_text(text):
    ents[1][1].delete(0, END)
    ents[1][1].insert(0, text)
    return


def keyfunction(x):
    '''
        used to sort mixed list of column names
    '''
    v = x
    if isinstance(v, int):
        v = '0%d' % v
    return v


def choose_default(i, collist, field):
    '''
        If default fields exist on the excel sheet,
        automatically choose them in the combobox values
    '''
    if field in ['streetnumber', 'streetname'] and 'originaladdress' in collist:
        combs[i][1].current(0)
    elif field in sorted(collist, key=keyfunction):
        location = [i for i,x in enumerate(sorted(collist, key=keyfunction)) if x == field][0]
        combs[i][1].current(location)
    else:
        combs[i][1].current(0)


def browsexlsx():
    from tkinter.filedialog import askopenfilename
    from os import path

    root1 = Tk()
    root1.withdraw()
    # possible other option: multifile=1
    # '.xls*' doesn't work on Mac.
    filenames = askopenfilename(parent=root, filetypes=[('Excel files', ['.xlsx','.xls']), 
                                                        ('Comma Separated Files', '.csv'),
                                                        ('All files', '.*')],
                                initialdir=path.dirname(r"Z:\EAD\DOL Data\QCEW to RPAD address merge\forgbat"))
    # response = root1.tk.splitlist(filenames)
    # for f in response:
    #     print(f)
    print(filenames)
    set_text(filenames)


def loadxlsx():
    '''
    Get the sheet names in the chosen excel file.
    '''
    filename = ents[1][1].get()
    if (filename.endswith(".xlsx") or filename.endswith(".xls") or filename.endswith(".csv")):
        # Read in the Load The Sheets.
        print("Processing: %s" % filename.encode().decode())
        f = path.basename(filename)
        status.set("Status: loading sheets of %s" % f.encode().decode())
        sheet_combo['state'] = 'enabled'
        frow['state'] = 'enabled'
        b3['state'] = 'enabled'
        output['state'] = 'enabled'
        if filename.endswith(".csv")==False:
            adds = ExcelFile(filename)
            print("Your Excel file includes these sheets: %s" % adds.sheet_names)
            sheet_combo['values'] = adds.sheet_names
        else:
            sheet_combo['values'] = ["Not Applicable: CSV File"]
        output.delete(0, END)
        output.insert(0, filename.replace(".csv", "_out.csv"))
        sheet_combo.current(0)
        status.set("Status: Choose which sheet of %s to process and press 'Load Fields'" % f.encode().decode())
    else:
        messagebox.showinfo(title="Not Excel or CSV File", message="Enter an Excel File. If it is a CSV file, make sure its extension is .csv")


def loadfields():
    '''
    Get the variable names in the chosen excel sheet
    '''

    filename = ents[1][1].get()
    f = path.basename(filename)
    status.set("Status: loading data and column names of %s" % f.encode().decode())
    
    #   if first row is not entered, assume 1 and set the form to 1.
    if frow.get() == "":
        frow.insert(0, 1)
        first_row = 1
    else:
        first_row = int(frow.get())
    if f.endswith(".csv")==False:
        adds = ExcelFile(filename)
        sheet = sheet_combo.get()
        print("%s and %s onwards chosen." % (sheet, first_row))
        df = adds.parse(sheet, skiprows=first_row - 1)
    else:
        df = pd.read_csv(filename, header=first_row - 1, low_memory=False)
    # print(df.columns.values)
    print("There are %s observations on this file." % len(df.index))
    ['Business Name:', 'Street Number:',
     'Street Name:', 'City/Borough:', 'Zipcode:', 'Boro Code:']
    defaults = {0: 'trade',
                1: 'legal',
                2: 'originaladdress',
                3: 'streetnumber',
                4: 'streetname',
                5: 'Borough',
                6: 'pzip',
                7: 'boro',
                8: 'state'}
    for i in range(len(combos)):
        collist = list(df.columns.values)
        collist.append("")
        combs[i][1]['state'] = 'enabled'
        combs[i][1]['values'] = sorted(collist, key=keyfunction)
        choose_default(i, collist, defaults[i])
    chk['state'] = 'enabled'
    b4['state'] = 'enabled'
#    print(combs[0][0], df[combs[0][1].get()].head(10))
    status.set("Status: Choose address fields, optionally edit output file, and press 'Geocode'")
    global DFrame
    DFrame = df
    return df


def Geocode(df, combs):
    '''
        Description: Geocodes the records that were not found in
        GBAT using Google API.
        By default all the years are in the exclude list so that 
        we don't redo any years and don't pay Google twice.
    '''
    status.set("Status: began geocoding")
    google_api_key2 = ents[0][1].get()

    try:
        gmaps = Client(key=google_api_key2)
    except:
        try:
            config = RawConfigParser()
            config.read('./API_Keys.cfg')
            google_api_key = config.get('Google', 'QCEW_API_Key')
            gmaps = Client(key=google_api_key)
        except:
            messagebox.showinfo("Can't Connect to Google", "Oups! Please check \
that your API key is valid, doesn't have leading/trailing spaces, and you are \
connected to internet! \nYour API key looks like vNIXE0xscrmjlyV-12Nj_BvUPaw")
            return None

    # test
    try:
        geocode_result = gmaps.geocode("Indepenent Buget Office, \
New York, NY 10003")
        # print(geocode_result[0]['geometry']['location']['lat'],
        #      geocode_result[0]['geometry']['location']['lng'])
        print("Google Connection Test Completed successfully.")
    except:
        print("Google Connection Test Failed.")

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
        df['Generated_streetaddress'] = df[combs[3][1].get()] + " " + df[combs[4][1].get()]
    else:
        df['Generated_streetaddress'] = df[combs[2][1].get()]

    if combs[7][1].get() == '':
        if combs[5][1].get() == '':
            df['no_boro'] = 1
            vals = combs[7][1]['values'] + ('no_boro',)
            combs[7][1]['values'] = sorted(vals, key=keyfunction)
            choose_default(7, vals, 'no_boro')
            df['no_city'] = 'New York'
        else:
            df['no_city'] = df[combs[5][1].get()]
    else:
        try:
            to_numeric(df[combs[7][1].get()], errors='raise')
        except:
            messagebox.showinfo("Error: Non-numeric Boro Codes!", "Boro code is needs to be a number\
 between 1 and 5. Either choose a different field or no field at all. \n\nChoosing no field will \
 result in 'New York City' assumed for all addresses.")
            return None
        df['no_boro'] = df[combs[7][1].get()]
        df['no_city'] = df['no_boro'].map(Boros)

    if combs[6][1].get() == '':
        df['no_zip'] = ''
        vals = combs[7][1]['values'] + ('no_zip',)
        combs[6][1]['values'] = sorted(vals, key=keyfunction)
        choose_default(6, vals, 'no_zip')
    else:
        if issubdtype(df[combs[6][1].get()].dtype, number):
            df['no_zip'] = df[combs[6][1].get()].round(0)
        else:
            df['no_zip'] = df[combs[6][1].get()]

    if combs[8][1].get() == '':
        df['no_state'] = 'NY'
        vals = combs[7][1]['values'] + ('no_state',)
        combs[6][1]['values'] = sorted(vals, key=keyfunction)
        choose_default(6, vals, 'no_state')
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
    if (second_run_state.get() is True and
        (combs[0][1].get() != "" or
         combs[1][1].get() != "")):
        add_list = ['Address', 'NameAddress']
        # Create a dataframe with unique observations on add_list
        df_unique = df[['Address', 'NameAddress', 'Generated_streetaddress', 'Goog_ID']].copy()
        df_unique.drop_duplicates(subset=add_list, keep="first", inplace=True)
        df_unique.reset_index(inplace=True)

        obs = len(df_unique.index) * 2
        print('There are %s unique observations to process...' % (obs))
    else:
        add_list = ['Address']
        # Create a dataframe with unique observations on add_list
        df_unique = df[['Address', 'Generated_streetaddress', 'Goog_ID']].copy()
        df_unique.drop_duplicates(subset=add_list, keep="first", inplace=True)
        df_unique.reset_index(inplace=True)

        obs = len(df_unique.index)
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

    directory = path.dirname(output.get())
    count_query = 0
    for var in add_list:
        print('Started checking variable ', var)
        i += 1
        for index, row in df_unique.iterrows():
            # set index <= len(df_unique.index) to process all observations.
            if index <= len(df_unique.index) and not(var in ['Address'] and row['Generated_streetaddress'] == ''):
                geocode_result = gmaps.geocode(row[var])
                count_query += 1
                status.set("Status: looking up observation %s of %s" % (count_query, obs))
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
                if index % 500 ==0:
                    df_unique.to_csv(path.join(directory, "GOOGLE_recovery.csv").encode().decode())
                    print("Recovery File GOOGLE_recovery.csv Saved when index was %s at %s" % (index, datetime.now()))
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

    # Merge back unique addresses with geocodes with original df.
    result = merge(df, df_unique, on=add_list, how='outer')
    result.drop(['index_x', 'Goog_ID_x', 'Goog_ID_y'], axis=1, inplace=True)
    # Update Google output fields with new values if they already existed on the file.
    for col in ['Gformatted_address0', 'Glat0', 'Glon0', 'GPartial0', 'Gtypes0', 'Gformatted_address1',
                'Glat1', 'Glon1', 'GPartial1', 'Gtypes1', 'Borough0', 'Borough1', 'Gzip0', 'Gzip1',
                'Gnumber0', 'Gnumber1', 'Gstreet0', 'Gstreet1', 'Both_Run_Same']:
        # print(col)
        if (col + '_y' in result.columns.values) and (col + '_x' in result.columns.values):
            result[col] = result[col + '_y'].fillna(result[col + '_x'])
            result.drop([col + '_y', col + '_x'], axis=1, inplace=True)

       # Drop the temporary variables.
    try:
        result.drop(['Generated_streetaddress_x','Generated_streetaddress_y'], axis=1, inplace=True)
        result.drop(['Address_x', 'Address_y'], axis=1, inplace=True)
    except:
        None
    try:
        result.drop(['NameAddress'], axis=1, inplace=True)
        print('dropped')
    except:
        None

    # ExcelFile(output.get())
    try:
        outfile = output.get()
        if outfile.endswith(".csv") is False:
            writer = ExcelWriter(outfile)
            result.to_excel(writer, 'Geocoded')
            writer.save()
        else:
            if path.exists(outfile.decode().encode()) is False:
                result.to_csv(outfile)
        message = outfile + "\n was successfully saved!\n There were %s queries made to Google Maps API" % (count_query)
    except:
        if outfile.endswith(".csv") is False:
            new_outfile = path.join(directory, "Google_Geocoded_" +
                                           time.strftime("%Y%m%d-%H%M%S") +
                                           ".xlsx").encode().decode()
            writer = ExcelWriter()
            result.to_excel(writer, 'Geocoded')
            writer.save()
        else:
            new_outfile = path.join(directory, "Google_Geocoded_" +
                                           time.strftime("%Y%m%d-%H%M%S") +
                                           ".csv").encode().decode()
            result.to_csv(new_outfile)
        
        message = ("Couldn't write to " + output.get() + ".\n Saved " +
                    path.basename(new_outfile) + " to the same directory.\n There were %s queries made to Google Maps API" % (count_query))
#   remove the recovery file.

    remove(path.join(directory, "GOOGLE_recovery.csv").encode().decode())
    print('Processed data and saved: ', output.get())

    endTime = time.time()
    t = (endTime - startTime) / 60
    status.set('Status: Done. It took %s minutes to make %s queries.' % (round(t, 2), count_query))
    print('Took %s minutes to run.' % round(t, 2))

    messagebox.showinfo('Success!', message)


# Run the program.
if __name__ == '__main__':
    root = Tk()

    # Quick workaround: Window background is white on Mac while buttons ... are grey.
    #root.configure(background='grey91')

    root.title("Google Geocoding")
    ents, row = makeform(root, fields)
    root.bind('<Return>', (lambda event, e=ents: fetch(e)))
    b1 = Button(row, text='Browse...', command=browsexlsx)
    b1.pack(side=LEFT, padx=5, pady=5)
    b2 = Button(row, text='Load Sheets', command=loadxlsx)
    b2.pack(side=LEFT, padx=5, pady=5, anchor=W)
    row.pack(side=TOP, fill=X, padx=5, pady=5)



    row = Frame(root)
    lbl_sheet = Label(row, text="Choose Input Sheet:", width=20, anchor='w')
    sheet_combo = Combobox(row, width=20, state='disabled')
    lbl_frow = Label(row, text="First Row?", width=10, anchor='w')
    frow = Entry(row, width=5, state='disabled')
    lbl_sheet.pack(side=LEFT, padx=5, pady=5)
    sheet_combo.pack(side=LEFT, expand=YES, fill=X)
    lbl_frow.pack(side=LEFT, padx=5, pady=5)
    frow.pack(side=LEFT, expand=YES, fill=X)

    b3 = Button(row, text='Load Fields', command=loadfields, state='disabled')

    b3.pack(side=LEFT, padx=5, pady=5, anchor='w')
    row.pack(side=TOP, fill=X, padx=5, pady=5)

    combs, row2 = makecomboboxes(root, combos)

    row = Frame(root)
    second_run_state = BooleanVar()
    second_run_state.set(True) #set check state
    chk = Checkbutton(row, text="Run twice using Address and Trade & Address", var=second_run_state, state='disabled')
    chk.pack(side=LEFT, padx=15, pady=5, anchor='w')
    row.pack(side=TOP, fill=X, padx=5, pady=5)

    row = Frame(root)
    b4 = Button(row, text='Geocode', command=lambda: Geocode(DFrame, combs), state='disabled')
    out_lbl = Label(row, text="Output File", width=20, anchor='w')
    output = Entry(row, width=40, state='disabled')
    out_lbl.pack(side=LEFT, padx=5, pady=5)
    output.pack(side=LEFT, expand=YES, fill=X)
    b4.pack(side=LEFT, padx=5, pady=5, anchor='w')
    row.pack(side=TOP, fill=X, padx=5, pady=5)

    row = Frame(root)
    status = StringVar()
    status.set("Status: waiting for user input ...")
    status_bar = Label(row, textvariable=status, bo=0.1,
                       relief=SUNKEN, anchor='w')
    status_bar.pack(side=BOTTOM, fill=X, padx=5, pady=5)
    row.pack(side=TOP, fill=X, padx=5, pady=5)


    root.mainloop()
from logging import exception
import tkinter
from tkinter import messagebox, ttk, filedialog, StringVar, OptionMenu
from ttkthemes import ThemedTk
import pandas as pd
import os
from pathlib import Path
from win32com.client import Dispatch
import re
import locale
locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')
import numpy as np
from datetime import datetime


##########      TKINTER INITIATOR       ##########
root = ThemedTk(theme="breeze")
root.title('Welcome Letter v2.0')
try:
    root.iconbitmap(r"path to icon")
except:
    pass




##########      PROGRAM VARIABLES       ##########
headers = ['RLMS LOAN NO','OLD LOAN NUMBER','INVESTOR ID','MAN CODE','MORTGAGOR LAST NAME','MORTGAGOR FIRST NAME','MORTGAGOR MIDDLE NAME','CO MORTGAGOR LAST NAME',
            'CO MORTGAGOR FIRST NAME','CO MORTGAGOR MIDDLE NAME','PROPERTY STREET ADDRESS','CITY NAME','PROPERTY ALPHA STATE CODE',
            'PROPERTY ZIP CODE','BILLING ADDRESS LINE 3','BILLING ADDRESS LINE 4','BILLING CITY NAME','BILLING STATE','BILLING ZIP CODE','PRIOR SERVICER',
            'TRANSFER DATE','PRIOR SERVICER PHONE','PRIOR SERVICER HOURS','DAY BEFORE TRANSFER','AS OF DATE','TOTAL MONTHLY PAYMENT','NEXT PAYMENT DUE DATE',
            'CURRENT CREDITOR','VOD REQ','TEMP COUPON','ACH FORM','PRIOR SERVICER ADDRESS 1','PRIOR SERVICER ADDRESS 2','ACQ TYPE',
            'LETTER_TYPE','LETTER_DATE','Calc Day Delq','BILLING ADDRESS LINE 2','ACQUISITION ID','ELOC Indicator',
            'ELOC Flag','Last Statement Date','Total Debt Last Statement','Interest Charged','Late Fees Charged','Other Fees Charged','Fees Charged','Escrow Charged',
            'Credits','Total Debt Now','DVN Expiration Date', 'DSI FLAG','PLS ID','Ubt post 2']

two_dec_place_col = ['TOTAL MONTHLY PAYMENT','Total Debt Last Statement','Interest Charged','Late Fees Charged',
                    'Other Fees Charged','Fees Charged','Escrow Charged','Credits','Total Debt Now']

address_column_w_format_issues = ['PROPERTY STREET ADDRESS', 'BILLING ADDRESS LINE 4', 'PRIOR SERVICER ADDRESS 2', 'PRIOR SERVICER']

address_pattern_dict = {'\s[A-Za-z]{2}\s':lambda match: match.group(0).upper() , '^[A-Za-z]{2}\s':lambda match: match.group(0).upper(),
                '\d{1}Th\s{1}': lambda match: match.group(0).lower(), '[-]\s{1}': '-', '\sDR\s':' Dr ', '\d{1}Nd\s{1}':lambda match: match.group(0).lower(),
                '\s{1}RD\s':' Rd ', '\sST\s':' St ', '--':'-', '\sSe$':' SE', '\sSw$':' SW', '\sNe$':' NE', '\sNw$':' NW'} 

company_name_columns = ['MORTGAGOR LAST NAME','MORTGAGOR FIRST NAME','MORTGAGOR MIDDLE NAME', 'PRIOR SERVICER', 'CURRENT CREDITOR']

company_name_pattern_dict = {'\s{1}Llc$|^Llc$':lambda match: match.group(0).upper(), '\s{1}Iii|Iii':lambda match: match.group(0).upper()}

states_list = [ 'AK', 'AL', 'AR', 'AZ', 'CA', 'CO', 'CT', 'DC', 'DE', 'FL', 'GA',
            'HI', 'IA', 'ID', 'IL', 'IN', 'KS', 'KY', 'LA', 'MA', 'MD', 'ME',
            'MI', 'MN', 'MO', 'MS', 'MT', 'NC', 'ND', 'NE', 'NH', 'NJ', 'NM',
            'NV', 'NY', 'OH', 'OK', 'OR', 'PA', 'RI', 'SC', 'SD', 'TN', 'TX',
            'UT', 'VA', 'VT', 'WA', 'WI', 'WV', 'WY']

todays_date = datetime.today().strftime('%m%d%y')
downloads_path = str(Path.home() / "Downloads")
downloads_path = downloads_path.replace('\\\\', '\\')


##########      SUB FUNCTION        ##########
def clearall():
    PassportEntry.delete(0,'end')
    XREFEntry.delete(0,'end')
    submitButton["state"] = "enable"

def passport_browse_cmd():
    """Opens file explorer browse dialogue box for user to search for files in GUI."""
    PassportEntry.delete(0,'end')
    root.filename = filedialog.askopenfilename()
    PassportEntry.insert(0, root.filename)
    return None

def XREF_browse_cmd():
    """Opens file explorer browse dialogue box for user to search for files in GUI."""
    XREFEntry.delete(0,'end')
    root.filename = filedialog.askopenfilename()
    XREFEntry.insert(0, root.filename)
    return None

def checkIfDuplicates(listOfElems):
    if len(listOfElems) == len(set(listOfElems)):
        return False
    else:
        return True




##########      MAIN FUNCTION       ##########


def main_func():
    
                        #####    STEP 1: LOAD DATE AND PREFORM INITIAL CHECKS      #####

    submitButton["state"] = "disabled"
    passport_path, XREF_path = PassportEntry.get(), XREFEntry.get()
    
    if passport_path=="" or XREF_path=="":
        tkinter.messagebox.showerror('Empty File Path','Please select the Passport Welcome Letter Template and Cross Reference file before clicking submit.')
        clearall()
        return
    if os.path.splitext(passport_path)[1] != '.xlsx' or os.path.splitext(XREF_path)[1] != '.xlsx':
        tkinter.messagebox.showerror('Unsupported File Type','This program only supports cvs and excel files at this time.')
        clearall()
        return
    
    try:
        passport = pd.read_excel(passport_path, engine='openpyxl', dtype={'LOAN NUMBER':'str', 'OLD LOAN NUMBER':'str', 'INVESTOR ID':'str',
                                                                        'MORTGAGOR FIRST NAME':'str','MORTGAGOR MIDDLE NAME':'str','MORTGAGOR LAST NAME':'str','CO MORTGAGOR LAST NAME':'str',
                                                                        'CO MORTGAGOR FIRST NAME':'str','CO MORTGAGOR MIDDLE NAME':'str','PROPERTY STREET NUMBER':'str',
                                                                        'PROPERTY STREET DIRECTION':'str', 'PROPERTY STREET NAME':'str', 'PROPERTY ZIP CODE':'str','BILLING ZIP CODE':'str'})
        passport.dropna(axis=0, how='all', inplace=True)
    except exception as e:
        tkinter.messagebox.showerror('Template Error', f'An error occurred when reading the Passport Welcome letter file. Here is the error message: {e}')
        clearall()
        return
        
    try:
        XREF = pd.read_excel(XREF_path, engine='openpyxl', dtype={'RLMS Loan No':'str'})
        XREF.dropna(axis=0, how='all', inplace=True)
    except exception as e:
        tkinter.messagebox.showerror('XREF Error', f'An error occurred when reading the XREF file. Here is the error message: {e}')
        clearall()
        return
    
    if 'LOAN NUMBER' not in passport.columns[0] or 'RLMS Loan' not in XREF.columns[0]:
        tkinter.messagebox.showerror('File Verfication Error', 'Please make sure you selected the correct file for each entry and try again.')
        clearall()
        return
    
    if checkIfDuplicates(list(XREF['RLMS Loan No'])) == True:
        tkinter.messagebox.showerror('XREF Duplicate Loan', 'A duplicate loan was found in the XREF file. Please resolve this error and try again')
        clearall()
        return
    if checkIfDuplicates(list(passport['LOAN NUMBER'])) == True:
        tkinter.messagebox.showerror('Passport Duplicate Loan', 'A duplicate loan was found in the Welcome Letter Template file. Please resolve this error and try again')
        clearall()
        return
    
    s = set(passport['LOAN NUMBER'])
    missing_loans = [x for x in XREF['RLMS Loan No'] if x not in s]
    
    if len(missing_loans) != 0:
        tkinter.messagebox.showerror('Missing Loans', f' {len(missing_loans)} loans were missing from the Passport. \nA list of these loans will be saved in your downloads folder as MISSING_LOANS_ACQID_DATE.CSV')



                        #####       STEP 2: CREATE MERGED DATAFRAME FROM CROSS REFERENCE AND PASSPORT FILE      #####

    merged_df = pd.merge(left=passport, right=XREF, left_on='LOAN NUMBER',right_on='RLMS Loan No' )

    merged_df['PROPERTY STREET DIRECTION'] = merged_df['PROPERTY STREET DIRECTION'].replace(np.nan, '', regex=True)
    
    ### ACQ_ID IS USED FOR THE FILE NAME WHEN EXPORTING THE WELCOME LETTER TEMP (SEE CODE AT STEP 5) 
    try:
        ACQ_ID = merged_df['Acquisition ID'][0]
    except:
        tkinter.messagebox.showerror('ACQ ID Error', f'The program was unable to find the "Acquisition ID" column in the xref file. \nPlease make sure the xref file contains the "Acquisition ID" column and try again.')
        clearall()
        return
    
    
    
    
                        #####       STEP 3: CREATE WELCOME LETTER TEMPLATE WITH 1 TO 1 ASSIGNMENTS      #####

    wel_temp = pd.DataFrame(columns=headers)
    try:
        wel_temp['RLMS LOAN NO'] = merged_df['LOAN NUMBER']
        wel_temp['OLD LOAN NUMBER'] = merged_df['OLD LOAN NUMBER']
        wel_temp['INVESTOR ID'] = merged_df['INVESTOR ID']
        wel_temp['MAN CODE'] = merged_df['MAN CODE']
        wel_temp['MORTGAGOR LAST NAME'] = merged_df['MORTGAGOR LAST NAME'].str.title()
        wel_temp['MORTGAGOR FIRST NAME'] = merged_df['MORTGAGOR FIRST NAME'].str.title()
        wel_temp['MORTGAGOR MIDDLE NAME'] = merged_df['MORTGAGOR MIDDLE NAME'].str.title()
        wel_temp['CO MORTGAGOR LAST NAME'] = merged_df['CO MORTGAGOR LAST NAME'].apply(lambda x: x.title() if isinstance(x, str) else x)
        wel_temp['CO MORTGAGOR FIRST NAME'] = merged_df['CO MORTGAGOR FIRST NAME'].apply(lambda x: x.title() if isinstance(x, str) else x)
        wel_temp['CO MORTGAGOR MIDDLE NAME'] = merged_df['CO MORTGAGOR MIDDLE NAME'].apply(lambda x: x.title() if isinstance(x, str) else x)
        wel_temp['PROPERTY STREET ADDRESS'] = merged_df['PROPERTY STREET NUMBER'].map(str) + ' ' + merged_df['PROPERTY STREET DIRECTION'].map(str) + ' ' + merged_df['PROPERTY STREET NAME']#.str.replace('-','')
        wel_temp['PROPERTY STREET ADDRESS'] = wel_temp['PROPERTY STREET ADDRESS'].str.title()
        wel_temp['PROPERTY STREET ADDRESS'] = wel_temp['PROPERTY STREET ADDRESS'].apply(lambda x: re.sub(' +', ' ', x))
        wel_temp['CITY NAME'] = merged_df['CITY NAME'].str.title()
        wel_temp['PROPERTY ALPHA STATE CODE'] = merged_df['PROPERTY ALPHA STATE CODE']
        wel_temp['PROPERTY ZIP CODE'] = merged_df['PROPERTY ZIP CODE'].str[:5]
        wel_temp['BILLING ADDRESS LINE 3'] = merged_df['BILLING ADDRESS LINE 3'].apply(lambda x: x.title() if isinstance(x, str) else x)
        wel_temp['BILLING ADDRESS LINE 4'] = merged_df['BILLING ADDRESS LINE 4'].apply(lambda x: x.title() if isinstance(x, str) else x)
        wel_temp['BILLING CITY NAME'] = merged_df['BILLING CITY NAME'].str.title()
        wel_temp['BILLING STATE'] = merged_df['BILLING STATE']
        wel_temp['BILLING ZIP CODE'] = merged_df['BILLING ZIP CODE'].str[:5]
        wel_temp['PRIOR SERVICER'] = merged_df['Servicer Name'].str.title()
        try:
            wel_temp['TRANSFER DATE'] = merged_df['Acquisition Date']
            wel_temp['DAY BEFORE TRANSFER'] = merged_df['Acquisition Date'].apply(pd.to_datetime) - pd.DateOffset(1)
            wel_temp['AS OF DATE'] = merged_df['Acquisition Date']
        except:
            wel_temp['TRANSFER DATE'] = merged_df['Acquisitions Date']
            wel_temp['DAY BEFORE TRANSFER'] = merged_df['Acquisitions Date'].apply(pd.to_datetime) - pd.DateOffset(1)
            wel_temp['AS OF DATE'] = merged_df['Acquisitions Date']
        wel_temp['PRIOR SERVICER PHONE'] = merged_df['Servicer Phone']
        wel_temp['PRIOR SERVICER HOURS'] = merged_df['Servicer Hours']
        wel_temp['TOTAL MONTHLY PAYMENT'] = merged_df['TOTAL MONTHLY PAYMENT'].astype(float)
        wel_temp['NEXT PAYMENT DUE DATE'] = merged_df['NEXT PAYMENT DUE DATE']
        wel_temp['CURRENT CREDITOR'] = merged_df['Current Creditor Name'].str.title()
        wel_temp['TEMP COUPON'] = 'N'
        wel_temp['ACH FORM'] =  'N'   
        wel_temp['PRIOR SERVICER ADDRESS 1'] = merged_df['Servicer Address'].str.title()
        wel_temp['PRIOR SERVICER ADDRESS 2'] = merged_df['Servicer City State Zip'].str.title()
        wel_temp['ACQ TYPE'] = 'M'    
        wel_temp['LETTER_TYPE'] = 'COMB'
        wel_temp['LETTER_DATE'] = pd.to_datetime("today") 
        wel_temp['Calc Day Delq'] =  wel_temp['NEXT PAYMENT DUE DATE'].apply(pd.to_datetime) - wel_temp['TRANSFER DATE'].apply(pd.to_datetime)
        wel_temp['ACQUISITION ID'] = merged_df['Acquisition ID']
        wel_temp['ELOC Indicator'] = merged_df['ELOC INDICATOR']
        wel_temp['ELOC Flag'] = ['Mortgage Loan' if x=='N' else 'Home Equity Line of Credit' for x in wel_temp['ELOC Indicator']]
        wel_temp['Last Statement Date'] = merged_df['DV BILLING STATEMENT DATE']
        wel_temp['Total Debt Last Statement'] = merged_df['DV BILL TOTAL DEBT AMOUNT'].astype(float)
        wel_temp['Interest Charged'] = merged_df['DV BILL INTEREST DUE AMOUNT'] - merged_df['DV NOTICE INTEREST BALANCE'].astype(float)
        wel_temp['Late Fees Charged'] = merged_df['DV BILL LATE CHARGE DUE AMOUNT'].astype(float) - merged_df['DV NOTICE LATE CHARGE AMOUNT'].astype(float)
        wel_temp['Other Fees Charged'] = merged_df['DV BILL OTHER FEES DUE AMOUNT'].astype(float) - merged_df['DV NOTICE OTHER FEES AMOUNT'].astype(float)
        wel_temp['Fees Charged'] = (merged_df['DV BILL CORP ADVANCE BALANCE'] - merged_df['DV NOTICE CORP ADVANCE AMOUNT']) + (merged_df['DV BILL LATE CHARGE DUE AMOUNT'] - merged_df['DV NOTICE LATE CHARGE AMOUNT']) + (merged_df['DV BILL OTHER FEES DUE AMOUNT'] - merged_df['DV NOTICE OTHER FEES AMOUNT'])
        wel_temp['Fees Charged'] = wel_temp['Fees Charged'].astype(float)
        wel_temp['Escrow Charged'] = merged_df['DV BILL ESCROW ADVANCE AMOUNT'].astype(float) - merged_df['DV NOTICE ESCROW ADV AMOUNT'].astype(float)
        wel_temp['Credits'] = merged_df['DV DIFFERENCE CREDITS AMOUNT'].astype(float)
        wel_temp['Total Debt Now'] = merged_df['DV NOTICE TOTAL DEBT AMOUNT'].astype(float)
        wel_temp['DVN Expiration Date'] = merged_df['DV DEBT VAL NOTICE END DATE']
        wel_temp['DSI FLAG'] = ['Y' if x=='C' or x=='D' else 'N' for x in merged_df['INTEREST CALC OPTION CODE']]
        wel_temp['PLS ID'] = ""
        wel_temp['Ubt post 2'] = ""
    except exception as e:
        tkinter.messagebox.showerror('Template Creation Error', f'An error occurred. Please inform the developer. \n Error: {e}')
        clearall()
        return





                        #####       STEP 4: APPLY FORMATTING TO MEET BUSINESS NEEDS     #####

    ###     RECREATES WLECOME LETTER TEMPLATE WITHOUT REO LOANS (MAN CODE "R")
    wel_temp = wel_temp[wel_temp['MAN CODE']!='R']

    ###   CHANGE CALC DAY DELQ TO INTEGERwel_temp['ELOC Indicator'] = passport['ELOC Indicator']
    wel_temp['Calc Day Delq'] = pd.to_numeric(wel_temp['Calc Day Delq'].dt.days, downcast='integer')

    ###   CONDITIONAL LOOP FOR  wel_temp['VOD REQ'] --------UNSTABLE 
    # for index, (man, state, delq) in enumerate(zip(wel_temp['MAN CODE'], wel_temp['PROPERTY ALPHA STATE CODE'],wel_temp['Calc Day Delq'])):
    #     if man in ['B', 'R']:
    #         wel_temp['VOD REQ'][index] = 'N'
    #         continue
    #     elif state in ['CA','WA','IL']:
    #         wel_temp['VOD REQ'][index] = 'Y'
    #         continue
    #     elif delq < -32:
    #         wel_temp['VOD REQ'][index] = 'Y'
    #         continue
    #     else:
    #         wel_temp['VOD REQ'][index] = 'N'
    
    ###     CONDITIONAL VOD REQ ---- STABLE (DOES THE SAME AS CODE ABOVE)
    wel_temp.loc[(wel_temp['MAN CODE'].isin(['B', 'R'])) & (~wel_temp['VOD REQ'].isin(['Y','N'])), 'VOD REQ'] = 'N'
    wel_temp.loc[(wel_temp['PROPERTY ALPHA STATE CODE'].isin(['CA','WA','IL'])) & (~wel_temp['VOD REQ'].isin(['Y','N'])), 'VOD REQ']='Y'
    wel_temp.loc[(wel_temp['Calc Day Delq']<-32) & (~wel_temp['VOD REQ'].isin(['Y','N'])), 'VOD REQ'] = 'Y'
    wel_temp.loc[~wel_temp['VOD REQ'].isin(['Y','N']), 'VOD REQ'] = 'N'

    ###     CHANGE $ VALUES TO 0.00 IF VOD REQ = N
    currency_columns = ['Total Debt Last Statement','Interest Charged','Late Fees Charged','Other Fees Charged','Fees Charged','Escrow Charged','Credits','Total Debt Now']
    wel_temp.loc[wel_temp['VOD REQ']=='N', currency_columns] = 0.00, 0.00,0.00, 0.00,0.00, 0.00,0.00, 0.00

    ###     CHANGES STATE AND ZIPCODE IF BILLING STATE NOT IN STATE LIST
    wel_temp.loc[~wel_temp['BILLING STATE'].isin(states_list), ['BILLING STATE','BILLING ZIP CODE']] = '', '00000'
    
    ###    FORMATTING ALL DATE COLUMNS (CONVERTS DATES TO MM/DD/YYYY)
    for i, value in wel_temp.dtypes.items():
        if 'datetime' in str(value):
            wel_temp[i] = wel_temp[i].dt.strftime('%m/%d/%Y')

    ###   FORMAT COLUMNS THAT REQUIRE 2 DECIMAL PLACES
    for col_header in two_dec_place_col:
        wel_temp[col_header] = [locale.format_string('%.2f', n, True) for n in wel_temp[col_header]]
    
    ###     REPLACE ALL NAN VALUES WITH AN EMPTY STRING  
    '''THIS IS REQUIRED FOR THE LOOPS BELOW'''
    wel_temp = wel_temp.replace(np.nan, '', regex=True)
    
    ###     LOOP THROUGH ADDRESS COLUMNS AND REPLACE/FORMAT VALUES IF THEY MATCH REGEX CRITERIA
    for col_header in address_column_w_format_issues:
        for key, value in address_pattern_dict.items():
            wel_temp[col_header] = [re.sub(key, value, x) for x in wel_temp[col_header]]
    
    ###     LOOP THROUGH NAME COLUMNS WITH COMPANY NAMES AND REPLACE/FORMAT VALUES IF THEY MATCH REGEX CRITERIA
    for col_header in company_name_columns:
        for key, value in company_name_pattern_dict.items():
            wel_temp[col_header] = [re.sub(key, value, x) for x in wel_temp[col_header]]

    
    
    
    
    
                        #####       STEP 5: EXPORT FINAL PRODUCT TO DOWNLOAD FOLDER AND OPEN IN EXCEL       #####
    
    wel_temp.to_excel(downloads_path + f'\COMBO WELCOME VOD LETTER TEMPLATE_{ACQ_ID}_{todays_date}.xlsx', index=False)
    
    xl = Dispatch("Excel.Application")
    xl.Visible = True # otherwise excel is hidden

    # newest excel does not accept forward slash in path
    wb = xl.Workbooks.Open(downloads_path +f'\COMBO WELCOME VOD LETTER TEMPLATE_{ACQ_ID}_{todays_date}.xlsx')
    
    if len(missing_loans)>0:
        try:
            missing_loans_df = pd.DataFrame(missing_loans)
            missing_loans_df.to_csv(downloads_path + f'\MISSING_LOANS_{ACQ_ID}_{todays_date}.csv', index=False)
        except exception as e:
            tkinter.messagebox.showerror('Missing Loans Export Error', f'An error occurred while exporting the missing loans. Please inform the developer. \n Error: {e}')
    
    
    clearall()
    submitButton["state"] = "enable"







##########      GUI DESIGN      ##########
PassportLabel = ttk.Label(root, text="Browse and select the LLB WELCOME LETTER TEMPLATE.")
PassportLabel.grid(row=0, column=0, pady=10, padx=10)

PassportEntry = ttk.Entry(root, width=50 )
PassportEntry.grid(row=1, column=0, pady=10, padx=10)

PassportbrowseButton = ttk.Button(root, text='Browse', command= passport_browse_cmd)
PassportbrowseButton.grid(row=1, column=1, pady=10, padx=10)

XREFLabel = ttk.Label(root, text="Browse and select the XREF file.")
XREFLabel.grid(row=4, column=0, pady=10, padx=10)

XREFEntry = ttk.Entry(root, width=50 )
XREFEntry.grid(row=5, column=0, pady=10, padx=10)

XREFbrowseButton = ttk.Button(root, text='Browse', command= XREF_browse_cmd)
XREFbrowseButton.grid(row=5, column=1, pady=10, padx=10)

submitButton = ttk.Button(root, text='Submit', command=main_func,  width=20)
submitButton.grid(row=6, column=0, padx=10, pady=10)

root.mainloop()

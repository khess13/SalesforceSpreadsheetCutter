import pandas as pd
import datetime as dt
import os

'''
Project built to cut 1 spreadsheet into smaller files for each corresponding
account in Salesforce.

Code expects additional input to properly id Accounts in Salesforce.

Code builds a file I'm calling a manifest which is used by Data Loader:
- creates a file's file
- associates file with Account
'''

#root will get current working directory
root = os.getcwd()
account_loc = root + '\\extract.csv'
datestamp = str(dt.datetime.now().strftime('%m-%d-%Y'))

#gathers files in root directory and returns only xlsx files
def get_files_from_dir(filepath, ext = '.XLSX'):
    filesindir = os.listdir(filepath)
    #tilda indicates open temp file, excluding these
    xlsxfiles = [f for f in filesindir if ext in f and not '~' in f]
    if len(xlsxfiles) == 0:
        print('No files found, try checking the extension.')
    else:
        return xlsxfiles

#get all xlsx in root
xlsx = get_files_from_dir(root)
#set up format of manifest for ContentVersion
contentVersion = pd.DataFrame(columns = ['Title','Description','VersionData',\
                                'PathOnClient','FirstPublishLocationId'])
#get account IDs by SCEIS code from Salesforce csv
accountids = pd.read_csv(account_loc)
#build dictionary because i don't know how to do this right
acctid_dict = {}
for index, row in accountids.iterrows():
    acctid_dict[row['CODE__C']] = row['ID']

print('Gathering outputs to parse.')
for x in xlsx:
    #open file, put in DataFrame
    xdf = pd.read_excel(x)
    #get rid of null Saless Contract
    xdf.dropna(subset = ['Sales Contract#'], inplace = True)
    #drops Customer accts with numbers
    #tests for numbers in customer, tilda is reversing
    xdf = xdf[~xdf['Customer'].str.isnumeric()].copy()
    #discover all agencies included in file
    xdf['AgyCode'] = xdf['Customer'].apply(lambda x: x[:4])
    agycodes = xdf['AgyCode'].drop_duplicates().tolist()


    #cut spreadsheets by agycode
    for agy in agycodes:
        #create subset of original data
        subdf = xdf[xdf['AgyCode'] == agy].copy()
        #determine total number of posting dates in file
        postingdaterange = xdf['Posting Date'].drop_duplicates().tolist()

        for date in postingdaterange:
            #make files idenifiers
            agycode = agy
            pdate = date.strftime('%m-%d-%Y')
            titledate = date.strftime('%Y-%m-%d') + ' Invoice'
            gendate = datestamp
            filename = agycode + ' Invoice Date ' + pdate + ' Generated On ' +\
                        gendate + '.xlsx'
            desc = 'Billing for services on ' + pdate + '. Generated on ' +\
                    gendate
            outputpath = root + '\\Cut Files\\'

            #gets Salesforce ID for account
            idofaccount = acctid_dict[agycode]
            #sub on date
            subsubdf = subdf[subdf['Posting Date'] == date].copy()

            #generating ContentVersion manifest
            nextentry = pd.Series([titledate, desc, outputpath + filename, \
                                    outputpath + filename, idofaccount], \
                                    index = contentVersion.columns)
            contentVersion = contentVersion.append(nextentry, ignore_index = True)

            #export file to excel file and save
            with pd.ExcelWriter(outputpath + filename) as writer:
                subsubdf.to_excel(writer, index = False)
            print('Creating ' + filename)

print('Creating manifest for ContentVersion')
contentVersion.to_csv(outputpath + 'ContentVersion Generated On ' + datestamp +\
                        '.csv', index = False)

print('Operation Complete!')

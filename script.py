import sys
import os
import comtypes.client
from mailmerge import MailMerge
import pandas as pd

os.chdir('C:\\Users\joefstedahl\PycharmProjects\Output')
'''
def docxToPDF():
    #i = 0
    wdFormatPDF = 17
    word = comtypes.client.CreateObject('Word.Application')
    #os.chdir('C:\\Users\joefstedahl\PycharmProjects\Output')
    #OutputFolder = os.listdir()
    #OutputPDF = '.\doc2text_{}'.format(i)
    #for files in OutputFolder:
    in_file = os.path.abspath('C:\\Users\joefstedahl\PycharmProjects\Output\GHRLN Company Listing.docx')
    out_file = os.path.abspath('C:\\Users\joefstedahl\PycharmProjects\Output\GHRLN Company Listing.pdf')
    pdf = word.Documents.Open(in_file)
    pdf.SaveAs(os.path.abspath(out_file), FileFormat=wdFormatPDF)
    pdf.Close()
    word.Quit()
docxToPDF()
'''

path = os.getcwd()

def process_CompanyListing(x):
    data_list = []
    #x[0].sort_values(['Contact Name: Account Name'], inplace=True)
    for row in x:
        account_name = row['Contact Name: Account Name']
        title = row['Contact Name: Title']
        # print (account_name)
        # print (title)
        data_list.append({
            'Contact_Name_Account_Name': account_name,
            'Contact_Name_Title': title,
        })
    #print(data_list)
    document.merge_rows('Contact_Name_Account_Name', data_list)
    os.chdir('C:\\Users\joefstedahl\PycharmProjects\Output')
    document.write(templates)
    os.chdir('C:\\Users\joefstedahl\PycharmProjects\TempsCompList')

'''
def process_MembershipListing(x):
    data_list = []
    for row in x:
        account_name = row['Contact Name: Account Name']
        title = row['Contact Name: Title']
        salutation = row['Contact_Name_Salutation']
        last_name = row[2]
        first_name = row[3]
        middle_ini = row[4]
        pref_name = row[5]
        title = row[6]
        data_list.append({
            'Contact_Name_Account_Name': account_name,
            'Contact_Name_Title': title,
            'Contact_Name_Salutation': salutation,
            'Contact_Name_First_Name': first_name,
            'Contact_Name_Last_Name': last_name,
            'Contact_Name_Preferred_Name': pref_name,
            'Contact_Name_Middle_Initial_Name': middle_ini,
        })
    document.merge_rows('Contact_Name_Account_Name', data_list)
    os.chdir('C:\\Users\joefstedahl\PycharmProjects\Output')
    document.write(templates)
    os.chdir('C:\\Users\joefstedahl\PycharmProjects\TempsCompList')
'''

os.chdir('C:\\Users\joefstedahl\PycharmProjects\SFExport')

df = pd.read_csv('src.csv', encoding='utf8')
os.chdir(path)

GHRLNdict = df[(df['Member Record: Network Name'] == 'GHRLN')].to_dict('record')
GSRNdict = df[(df['Member Record: Network Name'] == 'GSRN')].to_dict('record')
GTINdict = df[(df['Member Record: Network Name'] == 'GTIN')].to_dict('record')
GTRNdict = df[(df['Member Record: Network Name'] == 'GTRN')].to_dict('record')
HRINAsiadict = df[(df['Member Record: Network Name'] == 'HRIN-Asia')].to_dict('record')
HRRGdict = df[(df['Member Record: Network Name'] == 'HRRG')].to_dict('record')
IRNdict = df[(df['Member Record: Network Name'] == 'IRN')].to_dict('record')
IRNHCAdict = df[(df['Member Record: Network Name'] == 'IRN-HCA')].to_dict('record')

source = [GHRLNdict, GSRNdict, GTINdict, GTRNdict, HRINAsiadict, HRRGdict ,IRNdict, IRNHCAdict]

os.chdir('C:\\Users\joefstedahl\PycharmProjects\TempsCompList')
listtemplates = os.listdir()
for templates, src in zip(listtemplates, source):
    sorted_source = sorted(src, key=lambda k: k['Contact Name: Account Name'])
    #print (templates)
    document = MailMerge(templates)
    #print (document.get_merge_fields())
    process_CompanyListing(sorted_source)
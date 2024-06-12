import pandas as pd
import os
import openpyxl
import xlrd
from openpyxl import load_workbook
import datetime
import win32com
import win32com.client

BO_FILE_PP=('business objects location')
BO_FILE_NPP=(r'business objects location')
ADOBE_REQUESTS=(r'workbook')
#doing index_col=none will make sure that there is not an additional column that python puts in to list the rows. what we have below is a row reference - this isnt a problem. 
df=pd.read_excel(BO_FILE_PP, sheet_name = 'Web', header = None, index_col = None, index_row = None)
df2=pd.read_excel(BO_FILE_NPP, sheet_name = 'Report1', header = None, index_col = None, index_row = None)
df3=pd.read_excel(ADOBE_REQUESTS, sheet_name = '2019 UVs,Conversion & Orders RB', header = None, index_col = None, index_row = None)
df4=pd.read_excel(ADOBE_REQUESTS, sheet_name = 'PP Orders RR', header = None, index_col = None, index_row = None)
df.head()
df.columns
df.drop([0], inplace = True, axis = 1)
df.head()
df.drop([0,1],inplace = True, axis = 0)
df.head()
df2.head()
df2.columns
df2.drop([0,3,8], inplace = True, axis = 1)
df2.head()
df2.drop([0,1,2,3],inplace = True, axis = 0)
df2.head()
df3.head()
df4.head()
path = r'my location'
#loading workbook
book = load_workbook(path)
#object which writes to the workbook
writer = pd.ExcelWriter(path, engine='openpyxl') 
#where object and loadworkbook are introduced
writer.book = book

#writing in all sheet names from the workbook
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

# # df.to_excel(writer, sheet_name= 'BO data (PP) TY', header= None, index=0, startcol=2, startrow=1)
# # df2.to_excel(writer, sheet_name= 'BO Data (NPP) TY', header= None, index=0, startcol=0, startrow=1)
# # df3.to_excel(writer, sheet_name= '2019 UVs,Conversion & Orders RB', header= None, index=0, startcol=0, startrow=0)
# # #df4.to_excel(writer, sheet_name= 'PP Orders RR', header= None, index=0, startcol=0, startrow=0)

# #writer.save(r'my location')
# df.to_excel(writer, sheet_name= 'BO data (PP) TY', header= None, index=0, startcol=2, startrow=1)
# df2.to_excel(writer, sheet_name= 'BO Data (NPP) TY', header= None, index=0, startcol=0, startrow=1)
# # df3.to_excel(writer, sheet_name= '2019 UVs,Conversion & Orders RB', header= None, index=0, startcol=0, startrow=0)
# # df4.to_excel(writer, sheet_name= 'PP Orders RR', header= None, index=0, startcol=0, startrow=0)
# DATE=datetime.datetime.now().strftime("%d"+"%m"+"%Y")
# book.save(r'my location')
# #writer.save() this would have just saved this file in the same name as which it was uploaded.
# # XLSX_PATH ='my path'
# # XLSX_FOLDER= str(get_date(True))

# # def get_files(XLSX_PATH, XLSX_FOLDER):
# # # Open outlook
# # outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
# # # use email address
# # email=outlook.Folders['my email']
# # # navigate folders to where reports stored. You will need to personalise this
# # adobe = email.Folders['Inbox'].Folders['KPI Adobe Requests']

# # #import all email folder messages and data into 'messages'
# # messages=adobe.Items

# # date = datetime.today().date()

# # adobe_dict = [{'subject': 'Excel Workbook (KPI Template WTD Adobe Requests.xlsx)'},
# #                {'filename': 'KPI Template WTD Adobe Requests.xlsx'}]
              
# # print()
# # for file in adobe_dict:
# #     for message in messages:
# #         subject = file['subject']
# #         filename = file['filename']
# #         if message.Subject == subject and message.SentOn.date() == date:
# #             print(subject +' email found')
# #             for att in message.attachments:
# #                 att.SaveAsFile(CSV_PATH+CSV_FOLDER+"\\"+filename)
# #                 print()
              
    


# # for file in adobe_dict:
# #     for message in messages:
# #         subject = file['subject']
# #         filename = file['filename']
# #         if message.Subject == subject and message.SentOn.date() == date:
# #             print(subject +' email found')
# #             for att in message.attachments:
# #                 print(att.FileName+' found')
# #                 att.SaveAsFile(CSV_PATH+CSV_FOLDER+"\\"+filename)
# #                 print(att.FileName+' copied')
# #                 print()

#coding: utf-8

import pymssql
import win32com.client

def getFormat(someValue):
    temp=str(type(someValue))
    if str(type(someValue))=="<type 'float'>" or str(type(someValue))=="<type 'int'>":
        return "%d"
    elif str(type(someValue))=="<type 'str'>" or str(type(someValue))=="<type 'unicode'>" or str(type(someValue))=="<type 'time'>":
        return "%s"
    else:
        return '****************'

#**************************************************************************************************
"""блок импорта данных из EXCEL-файла"""
Excel = win32com.client.Dispatch("Excel.Application")
w_fb = Excel.Workbooks.Open('C:\\Users\\damir\\.virtualenvs\\win32com_pymssql\\import_spo_skidka_doplata\\src\\docs\\import_spo_skidka-doplata.xls')
activeSheetFB = w_fb.WorkSheets(1)
  
# print activeSheetFB.Cells(1,27).value

listRow=[]
strFormat=' VALUES('
Row = 3
while True:
    for Col in range(1,28):
        currCell=activeSheetFB.Cells(Row, Col).value
        if currCell:
            if str(type(currCell))=="<type 'time'>":
                currCell=str(currCell)
            listRow+=[currCell]
        else:
            listRow+=['NULL']
        strFormat+=getFormat(currCell)+','   
    Row += 1
    if Row == 4 : break  # if flagBrak : break
strFormat=strFormat[:-1]+")"

print listRow
# print strFormat
# print "fdsfsdfd %s blablabla" % listRow[3]
  
activeSheetFB = w_fb.WorkSheets(1)
w_fb.Close()
     
Excel.Quit()
"""конец блока импорта данных из EXCEL-файла"""
#**************************************************************************************************


    
#
strWhereToPut="hotel,room,htplace,meal,datebeg,dateend,rqdatebeg,rqdateend,dateinfrom,datein,dateoutfrom,dateout,nights,adult,child,infant,nobedchild,discount,\"percent\",net,rspos,rspotype,minlentour,maxlentour,pernight,rqdaysfrom,rqdaysstill"
#**************************************************************************************************
"""блок экспорта данных в sql-server!"""

conn = pymssql.connect(host='DMRPC', user='sa', password='1', database='WINTOUR')
cur = conn.cursor()

all_str="INSERT INTO import_hotelpr("+strWhereToPut+")"+strFormat
print all_str

# cur.execute("INSERT INTO import_hotelpr(datebeg) VALUES('12/31/15')")

# cur.executemany(all_str, tuple(listRow))
cur.executemany(all_str,[tuple(listRow)])
# cur.execute("INSERT INTO import_hotelpr(hotel) VALUES(NULL)")
conn.commit()
 
# cur.execute("Select *  from persons")
# row = cur.fetchone()
#  
# while row:
#     someStr = ''
#     for i in range(len(row)):
#         someStr+=str(row[i])+" "
#     print someStr
#     row = cur.fetchone()
# conn.commit()  
# conn.close()
"""конец блока экспорта данных в sql-server!"""
#**************************************************************************************************


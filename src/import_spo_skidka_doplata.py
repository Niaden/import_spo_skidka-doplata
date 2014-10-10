#coding: utf-8

import pymssql
# import win32com.client


conn = pymssql.connect(host='DMRPC', user='sa', password='1', database='WINTOUR')
cur = conn.cursor()
#cur.execute("alter table persons add purse int NULL ")
# for j in range(6):
#     cur.executemany("update persons set purse=%d where j=%d",j*100)
#cur.executemany("INSERT INTO persons VALUES(%d, %s, %d)", [ (1, 'John Doe', 200)])
# cur.execute("insert into persons values(300,'qqqq',10000) \
#                insert into persons values(400,'wwww',50000)")
cur.execute("Select *  from persons")
row = cur.fetchone()

while row:
    someStr = ''
    for i in range(len(row)):
        someStr+=str(row[i])+" "
    print someStr
    row = cur.fetchone()
conn.commit()  


conn.close()


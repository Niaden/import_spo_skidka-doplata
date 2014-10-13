#coding: utf-8


someStr="INSERT INTO import_hotelpr(\
hotel,room,htplace,meal,\
datebeg,dateend,rqdatebeg,rqdateend,dateinfrom,datein,dateoutfrom,dateout,\
nights,adult,child,infant,nobedchild,discount,\"percent\"\
,net,rspos,rspotype,\
minlentour,maxlentour,pernight,rqdaysfrom,rqdaysstill)\
 VALUES(%s,%s,%s,%s,\
 %s,%s,%s,%s,%s,%s,%s,%s,\
 %d,%d,%d,%d,%d,%s,%s,\
 %s,%s,%s,\
 %d,%d, \
 ,%s,%d,%d)" % (
                    u'ABBA XALET SUITES HOTEL "', u'STANDARD', u'DBL', u'BB',
                    '01/01/14 00:00:00', '12/31/15 00:00:00', '01/01/14 00:00:00', '12/31/15 00:00:00', 'NULL', 'NULL', 'NULL', 'NULL',
                    0, 20.0, 15.0, 10.0, 5.0, 'NULL', 'NULL',
                    u'\u0427\u0438\u0441\u0442\u043e\u0439', 'NULL', 'NULL',
                    0, 14.0,
                    u'\u0417\u0430 \u043f\u0435\u0440\u0438\u043e\u0434', 45.0, 60.0
                    )


print someStr

# 
#  
#  
#  
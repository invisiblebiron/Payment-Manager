# -*- coding: utf-8 -*-

import pymysql

#192.168.56.101
conn = pymysql.connect(host='192.168.7.88', 
                       port= 3306, 
                       user='root', 
                       #password='4xb@wickow@86', 
                       db='example')
cur = conn.cursor()
    
print(cur.execute("SHOW DATABASES"))

cur.close()
conn.close()
    
    #print('Datebase Version: ' + ver)
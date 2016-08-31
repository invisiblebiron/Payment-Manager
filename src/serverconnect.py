# -*- coding: utf-8 -*-

import pymysql
import sys
#192.168.56.101
conn = pymysql.connect(host='192.168.56.101', port=3306, user='biron', passwd='', db='testdb')
cur = conn.cursor()
    
ver = cur.fetchone()
    
    #print('Datebase Version: ' + ver)
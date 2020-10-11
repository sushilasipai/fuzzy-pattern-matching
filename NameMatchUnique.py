# -*- coding: utf-8 -*-
"""
Created on Wed Oct 16 10:46:05 2019

@author: sushila.sipai
"""

import cx_Oracle
import xlwt
import xlrd
from fuzzywuzzy import fuzz

ip='10.2.3.21'
port=1521
sid='DANPHE'
try:
    dns_tns=cx_Oracle.makedsn(ip,port,sid)
    db=cx_Oracle.connect('system','manager',dns_tns)
except:
    print('DB Connection Failed')
db.autocommit=True
    
namelist = []
book = xlwt.Workbook()
sh = book.add_sheet('Sheet 1')
col1_name = 'Original Name'
col2_name = 'Matched Name'
col3_name = 'Match Percentage'
sh.write(0,0,col1_name)
sh.write(0,1,col2_name)
sh.write(0,2,col3_name)

def closeMatches(uniquenames,word,index): 
    for matched in uniquenames:
        if(fuzz.ratio(word,matched)>94):
            sh.write(index,0,word)
            sh.write(index,1,matched)
            sh.write(index,2,fuzz.ratio(word,matched))
            index=index+1
            namelist.append(matched)
        
    return index;

         
if __name__ == "__main__": 
    wb=xlrd.open_workbook('test_9.xls')
    sheet=wb.sheet_by_index(0)
    wordlist=[sheet.cell_value(c,0) for c in range(sheet.nrows)]
    words=[words.encode('utf-8') for words in wordlist]
    
    cursor=db.cursor()
    cursor.execute("select acct_name from tbaadm.gam where acct_cls_flg='N' and acct_ownership<>'O'")
    patterns = [item[0] for item in cursor.fetchall()]
    print patterns
    uniquenames= list(set(patterns))
    index=1
    print uniquenames
    print "names printed"
    
    for word in words:
        print word
        new_index = closeMatches(uniquenames,word,index)
        index = new_index
        
book.save("names.xls")


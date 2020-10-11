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

def closeMatches(patterns,word,index): 
    for matched in patterns:
        if(fuzz.ratio(word,matched)>80):
            sh.write(index,0,word)
            sh.write(index,1,matched)
            sh.write(index,2,fuzz.ratio(word,matched))
            index=index+1
            namelist.append(matched)
        
    return index;

         
if __name__ == "__main__": 
    wb=xlrd.open_workbook('Book2.xls')
    sheet=wb.sheet_by_index(0)
    wordlist=[sheet.cell_value(c,0) for c in range(sheet.nrows)]
    words=[words.encode('utf-8') for words in wordlist]
    
    cursor=db.cursor()
    cursor.execute("select acct_name from tbaadm.gam where acct_cls_flg='N' and acct_ownership='C'")
    patterns = [item[0] for item in cursor.fetchall()]
    uniquenames= list(set(patterns))
    index=1
    
    for word in words:
        print word
        new_index = closeMatches(uniquenames,word,index)
        index = new_index
        
book.save("names.xls")

namelist = list(dict.fromkeys(namelist))

cur=db.cursor()

infoDetails = xlwt.Workbook()
sh = infoDetails.add_sheet('Sheet 1')
col1_name = 'Customer Name'
col2_name = 'Customer ID'
col3_name = 'Nationality'
col4_name = 'ID Document Type'
col5_name = 'Document Number'
col6_name = 'Father Name'
col7_name = 'Mother Name'
col8_name = 'Date of Birth'
col9_name = 'Risk Category'
col10_name = 'Risk Reason'

sh.write(0,0,col1_name)
sh.write(0,1,col2_name)
sh.write(0,2,col3_name)
sh.write(0,3,col4_name)
sh.write(0,4,col5_name)
sh.write(0,5,col6_name)
sh.write(0,6,col7_name)
sh.write(0,7,col8_name)
sh.write(0,8,col9_name)
sh.write(0,9,col10_name)

print("details from here")
ind=1;
for name in namelist:
    print name
    infoquery="""select distinct(cust_id),acct_name,country as nationality,uniqueidtype as id_document_type, 
                uniqueid as document_number,k.father_name,k.mother_name,to_char(cust_dob,'dd-mm-yyyy'),riskrating,
                manageropinion as risk_reason from tbaadm.gam left join crmuser.accounts
                on accounts.orgkey= gam.cif_id left join custom.kyc_details k on gam.cif_id= k.cif_id
                where acct_name = '""" +name +"' and gam.acct_cls_flg='N' and gam.acct_ownership='C'"

    cur.execute(infoquery)
    for curdata in cur:
        sh.write(ind,1,curdata[0])
        sh.write(ind,0,curdata[1])
        sh.write(ind,2,curdata[2])
        sh.write(ind,3,curdata[3])
        sh.write(ind,4,curdata[4])
        sh.write(ind,5,curdata[5])
        sh.write(ind,6,curdata[6])
        sh.write(ind,7,curdata[7])
        sh.write(ind,8,curdata[8])
        sh.write(ind,9,curdata[9])
        ind = ind+1

infoDetails.save('match_details.xls')     
  
        
        
        
        


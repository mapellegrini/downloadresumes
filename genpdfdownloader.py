#!/usr/bin/python 

import openpyxl

def name2url(namestr): 
  s1="https://twdc.sharepoint.com/sites/WDPR-fosrecruit/ProfInt/_layouts/15/download.aspx?SourceUrl=%2Fsites%2FWDPR-fosrecruit%2FProfInt%2FRecruiting%2F"
  s2="&FldUrl=&Source=https%3A%2F%2Ftwdc%2Esharepoint%2Ecom%2Fsites%2FWDPR%2Dfosrecruit%2FProfInt%2FRecruiting%2FForms%2FSU2014%5FCandidates%2Easpx"
  return s1 + namestr.replace(" ", "%20") + s2 

wb=openpyxl.load_workbook("/home/mark/Desktop/list.xlsx")
sheets=wb.sheetnames
sheet=wb[sheets[0]]

urls = [] 
a=2
while True: 
  mystr = sheet["A" + str(a)].value
  if (mystr==None):
    break 
  print mystr, name2url(mystr)  
  urls.append(name2url(mystr))   
  a+=1 

f=file("/tmp/downloadresumes.bat", "w") 
for url in urls:
  curstr = "certutil.exe -urlcache -split -f \"" + url + "\"\n"
  f.write(curstr) 
f.close() 

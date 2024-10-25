#!python3
#return.py - a test program for calculating automatic summation of the returned quantities across various jobs

import os,openpyxl,xlsxwriter,shutil
os.chdir('E:\\python\\desco item return calculation')
wb = openpyxl.load_workbook('return.xlsx')
a=wb.sheetnames
des={}
val={}
for i in a:
	sheet=wb[i]
	columns = list(sheet.iter_cols())
	c=columns[1]
	d=columns[2]
	t=columns[5]
	u=columns[6]
	for j in range(len(c)):
		if(c[j].value!=None and c[j].value!='Item Code'):
			des[c[j].value]=d[j].value
			if c[j].value not in val.keys():
				val[c[j].value]=0
			if t[j].value!=None and u[j].value!=None:
				val[c[j].value]+=(float(t[j].value)-float(u[j].value))
			'''
			if c[j].value not in val.keys():
				val[c[j].value]=0
			val[c[j].value]+=float(t[j].value-u[j].value)
			'''
for i in des.keys():
	print(i)
	print(des[i])
	print(val[i])

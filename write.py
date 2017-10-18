# -*- coding: utf-8 -*-  
import xdrlib,sys  
import xlwt  
import xlrd


path1='C:\\fenxi.xlsx'
#打开文件
data1=xlrd.open_workbook(path1)

sheet1=data1.sheet_by_index(5)
rows1=sheet1.nrows

path2='C:\\787981_land.xls'
#打开文件
data2=xlrd.open_workbook(path2)

sheet2=data2.sheet_by_index(0)
rows2=sheet2.nrows


#新建一个excel文件  
file = xlwt.Workbook()  
#新建一个sheet  
table = file.add_sheet('info',cell_overwrite_ok=True)  
#写入数据table.write(行,列,value) 
cell=sheet1.col(0)[0].value
table.write(0,0,cell)
cell=sheet1.col(1)[0].value
table.write(0,1,cell)
n=0

for i in range(1,rows1):
	cell1=sheet1.col(0)[i].value
	cell2=sheet1.col(1)[i].value
	table.write(i,0,cell1)
	table.write(i,1,cell2)
	for j in range(n,rows2):
		cell11=sheet2.col(0)[j].value
		cell22=sheet2.col(1)[j].value
		if(cell1==cell11):
			n=j
			table.write(i,2,cell11)
			table.write(i,3,cell22)
			break
 
#table.write(0,0,'wangpeng')  
#保存文件  
file.save('file.xls')  
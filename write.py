# -*- coding: utf-8 -*-  
import xdrlib,sys  
import xlwt  
import xlrd


path1='C:\\fenxi.xlsx'
#���ļ�
data1=xlrd.open_workbook(path1)

sheet1=data1.sheet_by_index(5)
rows1=sheet1.nrows

path2='C:\\787981_land.xls'
#���ļ�
data2=xlrd.open_workbook(path2)

sheet2=data2.sheet_by_index(0)
rows2=sheet2.nrows


#�½�һ��excel�ļ�  
file = xlwt.Workbook()  
#�½�һ��sheet  
table = file.add_sheet('info',cell_overwrite_ok=True)  
#д������table.write(��,��,value) 
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
#�����ļ�  
file.save('file.xls')  
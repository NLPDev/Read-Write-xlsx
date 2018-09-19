import xlwt

import numpy

from openpyxl import Workbook

import xlrd

loc = ("aa.xlsx")

wbr = xlrd.open_workbook(loc)
sheet = wbr.sheet_by_index(0)


nr=sheet.nrows
nc=sheet.ncols

for i in range(nr-1):
	st=sheet.cell_value(i+1, 0)

	wb = Workbook()
	sh = wb.active
	sh.title = 'Sheet1'
	sh.append(["Table1"])
	sh.merge_cells('A1:B1')

	for j in range(nc-1):
		c2 = sh.cell(row=2, column=j+1)
		c2.value = sheet.cell_value(0, j+1)



	st_type=sheet.cell_value(i+1, 1)
	st_size=sheet.cell_value(i+1, 2)
	res_type=st_type.split(",")
	res_size=st_size.split(",")

	for 
	st=st+".xlsx"
	wb.save(st)








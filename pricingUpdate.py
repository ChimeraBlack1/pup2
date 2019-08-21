import xlrd

loc = ("ProdMAPP.xlsx")

#open workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

#for row 0 and column 0
myVar = sheet.cell_value(0,0)

for x in range(0,10):
  myVar = sheet.cell_value(x,0)
  if myVar == "Accessories":
    print('found em boiz')

  print(myVar)
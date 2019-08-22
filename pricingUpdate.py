import xlrd

loc = ("ProdMAPP.xlsx")

#open workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

#for row 0 and column 0
testInput = sheet.cell_value(0,0)
groupModels = False
modelList = []
accList = []
xlListEnd = 600


for x in range(0,xlListEnd):
  testInput = sheet.cell_value(x,0)
  # if we find input named main unit, that signals a group of models
  if testInput == "Main Unit":
    #print(sheet.cell_value(x-1, 0))
    groupModels = True

  # if we hit 'accessories' stop grouping the models
  if testInput == "Accessories":
    groupModels = False

  # if groupModels is active, put the testInput into the models array.
  if groupModels == True and testInput != "Main Unit" and testInput != '':
    productNumber = int(sheet.cell_value(x,0))
    name = str(sheet.cell_value(x,1))
    desc = str(sheet.cell_value(x+1,1))
    mapp = int(sheet.cell_value(x,2))
    rmapp = int(sheet.cell_value(x,3))
    rmapp2 = int(sheet.cell_value(x,4))
    msrp = int(sheet.cell_value(x,5))

    newModel = {
      "productNumber": productNumber,
      "name": name,
      "desc": desc,
      "mapp": mapp,
      "rmapp": rmapp,
      "rmapp2": rmapp2,
      "msrp": msrp,
    }

    modelList.append(newModel)

#print(accList)
#print(modelList[1]["productNumber"])
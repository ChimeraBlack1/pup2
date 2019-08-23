import xlrd
import xlwt

loc = ("ProdMAPP.xlsx")

#open workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

#write to workbook
wbt = xlwt.Workbook()
wbt = wbt.add_sheet('my worksheet')
wbt.write(0,0, "new data yo")

#for row 0 and column 0
testInput = sheet.cell_value(0,0)
groupModels = False
groupAcc = False
modelList = []
accList = []
xlListEnd = 900


for x in range(0, xlListEnd):
  testInput = sheet.cell_value(x,0)
  # if we find input named main unit, that signals a group of models
  if testInput == "Main Unit":
    groupModels = True
    groupAcc = False
    #TODO need to reset the modelList on new Model group

  # if we hit 'accessories' stop grouping the models
  if testInput == "Accessories":
    groupModels = False
    groupAcc = True
    #TODO need to reset the accList on new accessory group

  # group models together to attach related accessories.
  if groupModels == True and testInput != "Main Unit" and testInput != '':
    productNumber = str(sheet.cell_value(x,0))
    name = str(sheet.cell_value(x,1))
    desc = str(sheet.cell_value(x+1,1))
    try:
      mapp = int(sheet.cell_value(x,2))
      rmapp = int(sheet.cell_value(x,3))
      rmapp2 = int(sheet.cell_value(x,4))
      msrp = int(sheet.cell_value(x,5))
    except ValueError:
      continue
  
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


  # create a list of accessories related to attach to these models
  if groupAcc == True and testInput != "Accessories" and testInput != '':
    productNumber = str(sheet.cell_value(x,0))
    name = str(sheet.cell_value(x,1))
    desc = str(sheet.cell_value(x+1,1))
    try:
      mapp = int(sheet.cell_value(x,2))
      rmapp = int(sheet.cell_value(x,3))
      rmapp2 = int(sheet.cell_value(x,4))
      msrp = int(sheet.cell_value(x,5))
    except ValueError:
      continue

    newAcc = {
      "productNumber": productNumber,
      "name": name,
      "desc": desc,
      "mapp": mapp,
      "rmapp": rmapp,
      "rmapp2": rmapp2,
      "msrp": msrp,
    }

    accList.append(newAcc)

print("Model 1 " + modelList[0]["name"])
# print(modelList)
# print("Accessory 1 " + accList[0]["name"])

# for i in range(0, len(accList)):
#   print(accList[i]["productNumber"])

for i in range(0, len(modelList)):
  print(modelList[i]["productNumber"])

# print(accList[16]["productNumber"])
# print(accList[16]["name"])
# print(accList[16]["desc"])
# print(accList[16]["mapp"])
# print(accList[16]["rmapp"])
# print(accList[16]["rmapp2"])
# print(accList[16]["msrp"])

# print("break")

# print(accList[17]["productNumber"])
# print(accList[17]["name"])
# print(accList[17]["desc"])
# print(accList[17]["mapp"])
# print(accList[17]["rmapp"])
# print(accList[17]["rmapp2"])
# print(accList[17]["msrp"])

# print("break2")

# print(accList[18]["productNumber"])
# print(accList[18]["name"])
# print(accList[18]["desc"])
# print(accList[18]["mapp"])
# print(accList[18]["rmapp"])
# print(accList[18]["rmapp2"])
# print(accList[18]["msrp"])


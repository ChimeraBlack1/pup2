import xlrd
import xlwt

loc = ("ProdMAPP.xlsx")

#open workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

#write to workbook
wbt = xlwt.Workbook()
wst = wbt.add_sheet('my worksheet')
# wst.write(0,0, "new data yo")
# wst.write(1,1,"test test test yo")
# wbt.save('example.xls')


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
    #if the accessory list is empty, that means we're still in the same model category
    if len(accList) <= 0:
      groupModels = True
      groupAcc = False
    else:
      # TODO need to reset the modelList on new Model group
      break

  # if we hit 'accessories' stop grouping the models
  if testInput == "Accessories":
    groupModels = False
    groupAcc = True
    #TODO need to reset the accList on new accessory group

  # group models together to attach related accessories.
  if groupModels == True and testInput != "Main Unit" and testInput != '':
    try:
      productNumber = int(sheet.cell_value(x,0))
    except:
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
    try:
      productNumber = int(sheet.cell_value(x,0))
    except:
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


print("total models: " + str(len(modelList)))

for i in range(0, len(modelList)):
  print(modelList[i]["name"] + " - " + str(modelList[i]["productNumber"]))

#add models found to CSV
for i in range(0, len(modelList)):
  wst.write(i, 0, "Model")
  wst.write(i, 3, "Ricoh Production MAPP")
  wst.write(i, 4, modelList[i]["name"])
  wst.write(i,5, "Equipment")
  wst.write(i,6, "Production")
  wst.write(i, 11, modelList[i]["productNumber"])
  wst.write(i, 13, modelList[i]["name"])
  wst.write(i,14, 0)
  wst.write(i,15, 0)
  wst.write(i,16, 0)
  wst.write(i,17, 0)
  wst.write(i, 18, modelList[i]["mapp"])
  wst.write(i, 19, modelList[i]["mapp"])
  wst.write(i, 20, modelList[i]["msrp"])
  wst.write(i, 29, modelList[i]["desc"])
  wst.write(i, 31, modelList[i]["rmapp"])
  wst.write(i, 32, modelList[i]["rmapp2"])
  for j in range(0,38):
    wst.write(i,33+j, 0)
  # TODO add all accessories related to this model here


wbt.save('UpdatedMAPP.xls')


# print("total accessories: " + str(len(accList)))
# for i in range(0, len(accList)):

#   print(accList[i]["name"] + " - " + accList[i]["productNumber"])


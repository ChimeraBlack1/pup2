import xlrd
import xlwt

loc = ("ProdMAPP.xlsx")

#open workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

#write to workbook
wbt = xlwt.Workbook()
wst = wbt.add_sheet('Updated Mapp')

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

  if testInput == "Service Data":
    groupModels = False
    groupAcc = False
    # TODO remove break for further logic
    break

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


# print("total models: " + str(len(modelList)))

# for i in range(0, len(modelList)):
#   print(modelList[i]["name"] + " - " + str(modelList[i]["productNumber"]))
modelStart = 0
accStart = 0
#add models found to XLS
for i in range(0, len(modelList)):
  print("accList: " + str(len(accList)))
  print("modelStart: " + str(modelStart))
  print("modelStart + accList length = " + str(len(accList) + modelStart))
  wst.write(modelStart, 0, "Model")
  wst.write(modelStart, 3, "Ricoh Production MAPP")
  wst.write(modelStart, 4, modelList[i]["name"])
  wst.write(modelStart, 5, "Equipment")
  wst.write(modelStart, 6, "Production")
  wst.write(modelStart, 11, modelList[i]["productNumber"])
  wst.write(modelStart, 13, modelList[i]["name"])
  wst.write(modelStart, 14, 0)
  wst.write(modelStart, 15, 0)
  wst.write(modelStart, 16, 0)
  wst.write(modelStart, 17, 0)
  wst.write(modelStart, 18, modelList[i]["mapp"])
  wst.write(modelStart, 19, modelList[i]["mapp"])
  wst.write(modelStart, 20, modelList[i]["msrp"])
  wst.write(modelStart, 29, modelList[i]["desc"])
  wst.write(modelStart, 31, modelList[i]["rmapp"])
  wst.write(modelStart, 32, modelList[i]["rmapp2"])
  
  # print("accList again: " + str(len(accList)))
  # print("modelStart again: " + str(modelStart))
  # print("modelStart again + accList length again = " + str(len(accList) + modelStart))

  accStart = accStart + 1
  #write in accessories
  for k in range(0, len(accList)):
    wst.write(accStart, 0, "Accessory")
    wst.write(accStart, 3, "Ricoh Production MAPP")
    wst.write(accStart, 4, accList[k]["name"])
    wst.write(accStart, 13, accList[k]["name"])
    accStart = accStart + 1
  
  modelStart = modelStart + len(accList) + 1
  
  
  # write a bunch of zeros in the special pricing fields
  for j in range(0,38):
    wst.write(modelStart,33+j, 0)
  # TODO add all accessories related to this model here
  # for k in range(1,len(accList)):
  #   wst.write(k, 0, "Accessory")
  # wst.write(i+1,0, "Accessory")
  # i = i + len(accList) + 1
  
  

wbt.save('UpdatedMAPP.xls')


# print("total accessories: " + str(len(accList)))
# for i in range(0, len(accList)):
#   print(accList[i]["name"] + " - " + str(accList[i]["productNumber"]))


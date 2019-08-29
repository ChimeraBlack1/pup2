import xlrd
import xlwt

loc = ("ProdMAPP.xlsx")

#open workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

#write to workbook
wbt = xlwt.Workbook()
wst = wbt.add_sheet('Updated Mapp')

# establish starting points
groupModels = False
groupAcc = False
groupService = False
groupGlobal = False
modelList = []
accList = []
globalList = []
xlListEnd = 2900

for x in range(0, xlListEnd):
  testInput = sheet.cell_value(x,0)

  # if we find "Professional Services", that signals the group of global accessories
  if testInput == "Professional Services":
    groupGlobal = True
    groupModels = False
    groupAcc = False
    groupService = False

  # if we find input named main unit, that signals a group of models
  if testInput == "Main Unit":
    groupModels = True
    groupAcc = False
    groupService = False
    groupGlobal = False

    #if the accessory list is empty, that means we're still in the same model category
    if len(accList) <= 0:
      groupModels = True
      groupAcc = False
      groupService = False
      groupGlobal = False
    else:
      # TODO need to reset the accList and modelList on new Model group
      ### WRITE TO XLS METHOD ###
      modelStart = 0
      accStart = 0
      globalStart = 0
      #add models found to XLS
      for i in range(0, len(modelList)):
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
        # write a bunch of zeros in the special pricing fields
        for j in range(0,38):
          wst.write(modelStart,33+j, 0)

        globalStart = globalStart + 1
        #write in accessories
        for k in range(0, len(globalList)):
          wst.write(globalStart, 0, "Access")
          wst.write(globalStart, 3, "Ricoh Production MAPP")
          wst.write(globalStart, 4, globalList[k]["name"])
          wst.write(globalStart, 8, "ACCESSORY")
          wst.write(globalStart, 9, "N")
          wst.write(globalStart, 11, globalList[k]["productNumber"])
          wst.write(globalStart, 13, globalList[k]["name"])
          wst.write(globalStart, 14, 0)
          wst.write(globalStart, 15, 0)
          wst.write(globalStart, 16, 0)
          wst.write(globalStart, 17, 0)
          wst.write(globalStart, 18, globalList[k]["mapp"])
          wst.write(globalStart, 19, globalList[k]["mapp"])
          wst.write(globalStart, 20, globalList[k]["msrp"])
          wst.write(globalStart, 25, 0)
          wst.write(globalStart, 29, globalList[k]["desc"])
          wst.write(globalStart, 31, globalList[k]["rmapp"])
          wst.write(globalStart, 32, globalList[k]["rmapp2"])
          # write a bunch of zeros in the special pricing fields
          for l in range(0,38):
            wst.write(globalStart,33+l, 0)

          globalStart = globalStart + 1

        accStart = globalStart
        #write in accessories
        for j in range(0, len(accList)):
          wst.write(accStart, 0, "Access")
          wst.write(accStart, 3, "Ricoh Production MAPP")
          wst.write(accStart, 4, accList[j]["name"])
          wst.write(accStart, 8, "ACCESSORY")
          wst.write(accStart, 9, "N")
          wst.write(accStart, 11, accList[j]["productNumber"])
          wst.write(accStart, 13, accList[j]["name"])
          wst.write(accStart, 14, 0)
          wst.write(accStart, 15, 0)
          wst.write(accStart, 16, 0)
          wst.write(accStart, 17, 0)
          wst.write(accStart, 18, accList[j]["mapp"])
          wst.write(accStart, 19, accList[j]["mapp"])
          wst.write(accStart, 20, accList[j]["msrp"])
          wst.write(accStart, 25, 0)
          wst.write(accStart, 29, accList[j]["desc"])
          wst.write(accStart, 31, accList[j]["rmapp"])
          wst.write(accStart, 32, accList[j]["rmapp2"])
          # write a bunch of zeros in the special pricing fields
          for m in range(0,38):
            wst.write(accStart,33+m, 0)

          accStart = accStart + 1
        
        globalStart = accStart
        modelStart = accStart
        
        ### /WRITE TO XLS METHOD ###
      print('end of config')
      accListLen = len(accList)
      modelListLen = len(modelList)
      print("Added: " + str(modelListLen) + " models")
      print("Attached " + str(accListLen) + " 'model specific' accesories to each model")
      modelList = []
      accList = []
      break

  # if we hit 'Accessories' stop grouping the models and start grouping the accessories
  if testInput == "Accessories":
    groupAcc = True
    groupModels = False
    groupService = False
    groupGlobal = False
    #TODO need to reset the accList on new accessory group
  
  # if we hit service data, GROUP SERVICE DATA
  if testInput == "Service Data":
    groupService = True
    groupModels = False
    groupAcc = False
    groupGlobal = False
    # TODO remove break for further logic
    print('hit service data')
  
  # GROUP MODELS
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


  # GROUP MODEL SPECIFIC ACCESSORIES
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

    if newAcc['name'] != 'NM':
      accList.append(newAcc)

  # GROUP GLOBAL ACCESSORIES
  if groupGlobal == True and testInput != "Professional Services" and testInput != '':
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

    globalList.append(newModel)


globalListLen = len(globalList)


print("Attached " + str(globalListLen) + " global accessories to each model")
print("totalling " + str((modelListLen + globalListLen + accListLen) * modelListLen) + " line items" )


wbt.save('UpdatedMAPP.xls')

# print("total globals: " + str(len(accList)))
# for i in range(0, len(accList)):
#   print(str(accList[i]["productNumber"]) + " - " + accList[i]["name"])


import xlrd
import xlwt
import math

def ReadNewMapp(newMappName):
  """
  Reads the new mapp that was created by pricingUpdate.py and scans product numbers.
  If a match is found it writes the model number in the new workbook
  """
  # read workbook
  loc = newMappName
  wb = xlrd.open_workbook(loc)
  #open workbook
  sheet = wb.sheet_by_index(0)

  # write to workbook
  wbt = xlwt.Workbook()
  wst = wbt.add_sheet('Updated Mapp 2')

  itemIndex = 1
  endOfNewMapp = 3000
  sherpaExportStart = 8
  sherpaExportEnd = 20

  for x in range(sherpaExportStart,sherpaExportEnd):
    productNumber = sheet.cell_value(x+8,11)
    assetID = sheet.cell_value(x+8,1)
    try:
      productNumber = int(productNumber)
    except:
      productNumber = str(productNumber)
    
    for y in range(0, endOfNewMapp):
      wst.write(x, 1, productNumber)
      wst.write(x, 2, assetID)
      itemIndex = itemIndex + 1
      wbt.save("IDAdded.xls")
      print(str(productNumber))


ReadNewMapp("SherpaDataExport.xlsm")

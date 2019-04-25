import openpyxl,datetime,csv,PIL,platesLoadsTransfer


excelSPR=r'C:\TEMP\Plate.xlsx'
templateFile=r'C:\TEMP\plateLoadsTemplate.xlsm'
loads=r'C:\TEMP\loads.csv'
excel=platesLoadsTransfer.Folder(excelSPR,templateFile)
print(excel.getSPRexcel())
print(excel.getDiameter())
excel.fillTemplate()
#loadFile=Load(loads,templateFile)
#print(loadFile.getLoads())
#print(loadFile.fillLoads())

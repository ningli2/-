import openpyxl,datetime,csv,PIL,platesLoadsTransfer


excelSPR=r'C:\TEMP\Copy of Plate.xlsx'
templateFile=r'C:\TEMP\plateInformation.xlsx'
loads=r'C:\TEMP\loads.csv'
#excel=Folder(excelSPR,templateFile)
#print(excel.getSPRexcel())
#print(excel.getDiameter())
#excel.fillTemplate()
loadFile=platesLoadsTransfer.Load(loads,templateFile)
print(loadFile.getLoads())
print(loadFile.fillLoads())

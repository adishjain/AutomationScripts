import xlrd
import xlwt
loc = ("C:/Users/Adish.Jain/Desktop/E DRIVE DATA/Engagements/HCL/March2019/Automation/test.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

removedBook = xlwt.Workbook()
removedSheet = removedBook.add_sheet('Removable')

draftReport = xlwt.Workbook()
detailedObservationsSheet = draftReport.add_sheet('DetailedObservations')

qidToBeRemoved = "C:/Users/Adish.Jain/Desktop/E DRIVE DATA/Engagements/HCL/March2019/Automation/test.txt"
QIDs = []
removedRows = []
retainedRows = []
with open(qidToBeRemoved) as fp:
    for cnt,line in enumerate(fp):
#        print("Line {}: {}".format(cnt,line))
        line = line.replace('\n','')
        QIDs.append(line)
#        print(QIDs)
count = 0
toBeRemoved=[]
toBeRetained=[]
totalCols = sheet.ncols
for i in range(sheet.nrows):
    val = sheet.cell_value(i,2)
#    print(type(val))
    val = str(val)
    val = val.replace('.0','')
#    print(type(val))
    if val in QIDs:
        toBeRemoved.append(i)
        row = sheet.row_values(i)
        removedRows.append(row)
        count = count + 1
    else:
        toBeRetained.append(i)
        row = sheet.row_values(i)
        retainedRows.append(row)
rowNum = 0

for r in removedRows:
    columnNum = 0
    for c in r:
        removedSheet.write(rowNum,columnNum,c)
        columnNum = columnNum + 1
    rowNum = rowNum + 1
removedBook.save("C:/Users/Adish.Jain/Desktop/E DRIVE DATA/Engagements/HCL/March2019/Automation/testRemoved.xls")

rowNum = 0
for r in retainedRows:
    columnNum = 0
    for c in r:
        detailedObservationsSheet.write(rowNum,columnNum,c)
        columnNum = columnNum + 1
    rowNum = rowNum + 1
draftReport.save("C:/Users/Adish.Jain/Desktop/E DRIVE DATA/Engagements/HCL/March2019/Automation/draftReport.xls")

#        for c in row:
#            removedSheet.write()
#            count = count+1
#           print(count,"  ",c)
#        print(count,"  ",row)

import xlrd
import XlsHelper
import Generator
rb = xlrd.open_workbook('source.xls', formatting_info=True)
sheet = rb.sheet_by_index(0)

xlsHelper = XlsHelper.XlsHelper(sheet)
columnNames = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y']
fileNameColumn = 'L'

generator = Generator.Generator('template.html', xlsHelper, columnNames, fileNameColumn)

startRow = 1
#count = 1
count = 1047
rowIndex = startRow

while rowIndex < count + startRow:
	try:
		generator.GenerateFor(rowIndex)
		rowIndex += 1
	except:
		print "row {0}".format(rowIndex)
		raise
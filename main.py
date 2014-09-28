import xlrd
import XlsHelper
from Generator import PluralGenerator

rb = xlrd.open_workbook('source.xls', formatting_info=True)
sheet = rb.sheet_by_index(0)

xlsHelper = XlsHelper.XlsHelper(sheet)
columnNames = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y']
fileNameColumn = 'L'

html_generator = PluralGenerator.PluralGenerator('template.html', xlsHelper, columnNames, fileNameColumn)

startRow = 1
rowIndex = startRow

try:
	while True:
		html_generator.generate_for(rowIndex)
		rowIndex += 1
except IndexError:
	print "generation completed"
except:
	print "row {0}".format(rowIndex)
	raise


import xlrd
import XlsHelper
from Generator import PluralGenerator, SingleGenerator


rb = xlrd.open_workbook('source.xls', formatting_info=True)
sheet = rb.sheet_by_index(0)

xlsHelper = XlsHelper.XlsHelper(sheet)
columnNames = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y']
fileNameColumn = 'L'

encode = 'windows-1251'
html_generator = PluralGenerator.PluralGenerator('template.html', xlsHelper, columnNames, encode, fileNameColumn, 'output')
table_generator = SingleGenerator.SingleGenerator('table_template.html', xlsHelper, columnNames, encode)
map_generator = SingleGenerator.SingleGenerator('map_template.xml', xlsHelper, columnNames, encode)

startRow = 1
rowIndex = startRow

try:
	while True:
		html_generator.generate_for(rowIndex)
		table_generator.generate_for(rowIndex)
		map_generator.generate_for(rowIndex)
		rowIndex += 1
except IndexError:
	print "generation completed"
except:
	print "row {0}".format(rowIndex)
	raise
finally:
	table_generator.save_results('table.html')
	map_generator.save_results('map.xml')

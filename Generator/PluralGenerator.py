import os
import Generator


class PluralGenerator(Generator):
	def __init__(self, templatePath, xlsHelper, columnNames, fileNameColumn, output_dir="output"):
		Generator.__init__(self, templatePath, xlsHelper, columnNames)
		assert fileNameColumn is not None

		self.fileNameColumn = fileNameColumn
		self.output_dir = output_dir

		# ensure that self.output_dir exists
		if not os.path.exists(self.output_dir):
			os.makedirs(self.output_dir)


	def GenerateFor(self, rowIndex):
		fileName = self.xlsHelper.GetValue(rowIndex, self.fileNameColumn)

		filePath = os.path.join(self.output_dir, fileName)
		paramMap = self.FetchMapFor(rowIndex)

		content = self.template % paramMap

		f = open(filePath, 'w')
		f.write(content.encode('windows-1251'))
		f.close()

	def FetchMapFor(self, rowIndex):
		paramMap = {}

		for colName in self.columnNames:
			v = self.xlsHelper.GetValue(rowIndex, colName)
			v = super.prepare_value(v)
			paramMap[colName] = v

		return paramMap

import os
class Generator:
	def __init__(self, templatePath, xlsHelper, columnNames, fileNameColumn, output_dir="output"):
		assert templatePath is not None
		assert xlsHelper is not None
		assert columnNames is not None
		assert columnNames is not None

		self.templatePath = templatePath
		self.xlsHelper = xlsHelper
		self.columnNames = columnNames
		self.fileNameColumn = fileNameColumn
		self.output_dir = output_dir

		# ensure that self.output_dir exists
		if not os.path.exists(self.output_dir):
			os.makedirs(self.output_dir)

		f = open(self.templatePath)
		self.template = f.read().decode('windows-1251')

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
			paramMap[colName] = self.xlsHelper.GetValue(rowIndex, colName)

		return paramMap

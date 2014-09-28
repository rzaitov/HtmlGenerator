import os
class Generator:
	def __init__(self, templatePath, xlsHelper, columnNames):
		assert templatePath is not None
		assert xlsHelper is not None
		assert columnNames is not None

		self.templatePath = templatePath
		self.xlsHelper = xlsHelper
		self.columnNames = columnNames

		f = open(self.templatePath)
		self.template = f.read().decode('windows-1251')


	# if value is float but can interpreted as int we will return int
	def prepare_value(self, raw_value):
		if isinstance(raw_value, float):
			if raw_value.is_integer():
				return int(raw_value)

		return raw_value


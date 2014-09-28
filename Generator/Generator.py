

class Generator(object):
	def __init__(self, template_path, xls_helper, column_names):
		assert template_path is not None
		assert xls_helper is not None
		assert column_names is not None

		self.template_path = template_path
		self.xls_helper = xls_helper
		self.column_names = column_names

		f = open(self.template_path)
		self.template = f.read().decode('windows-1251')

	# if value is float but can interpreted as int we will return int
	@staticmethod
	def prepare_value(raw_value):
		if isinstance(raw_value, float):
			if raw_value.is_integer():
				return int(raw_value)

		return raw_value


	def fetch_map_for(self, rowIndex):
		paramMap = {}

		for colName in self.column_names:
			v = self.xls_helper.GetValue(rowIndex, colName)
			v = self.prepare_value(v)
			paramMap[colName] = v

		return paramMap

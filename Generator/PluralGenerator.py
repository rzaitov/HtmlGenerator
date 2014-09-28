import os
import Generator


class PluralGenerator(Generator.Generator):
	def __init__(self, template_path, xls_helper, column_names, file_name_column, output_dir="output"):
		super(PluralGenerator, self).__init__(template_path, xls_helper, column_names)
		assert file_name_column is not None

		self.file_name_column = file_name_column
		self.output_dir = output_dir

		# ensure that self.output_dir exists
		if not os.path.exists(self.output_dir):
			os.makedirs(self.output_dir)

	def generate_for(self, row_index):
		file_name = self.xls_helper.GetValue(row_index, self.file_name_column )

		file_path = os.path.join(self.output_dir, file_name)
		param_map = self.fetch_map_for(row_index)

		content = self.template % param_map

		f = open(file_path, 'w')
		f.write(content.encode('windows-1251'))
		f.close()
import Generator


class SingleGenerator(Generator.Generator):
	def __init__(self, template_path, xls_helper, column_names):
		super(SingleGenerator, self).__init__(template_path, xls_helper, column_names)

		self.items = []
		self.start_loop_token = '<!--StartLoop-->'
		self.end_loop_token = '<!--EndLoop-->'

		self.start = self.template.find(self.start_loop_token)
		self.end = self.template.find(self.end_loop_token, self.start)

		self.loop_template_item = self.template[self.start + len(self.start_loop_token): self.end]
		self.loop_template_item = self.loop_template_item.lstrip()

	def generate_for(self, row_index):
		param_map = self.fetch_map_for(row_index)
		content = self.loop_template_item % param_map
		self.items.append(content)

	def save_results(self, output_file_name):
		assert output_file_name is not None

		f = open(output_file_name, 'w')
		for item in self.items:
			f.write(item.encode('windows-1251'))
		f.close()

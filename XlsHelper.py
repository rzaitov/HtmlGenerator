__author__ = 'rzaitov'
class XlsHelper:
	def __init__(self, sheet):
		assert sheet is not None

		self.sheet = sheet
		self.columNameIndexMap = {
			'A': self.GetColIndex('A'),
			'B': self.GetColIndex('B'),
			'C': self.GetColIndex('C'),
			'D': self.GetColIndex('D'),
			'E': self.GetColIndex('E'),
			'F': self.GetColIndex('F'),
			'G': self.GetColIndex('G'),
			'H': self.GetColIndex('H'),
			'I': self.GetColIndex('I'),
			'J': self.GetColIndex('J'),
			'K': self.GetColIndex('K'),
			'L': self.GetColIndex('L'),
			'M': self.GetColIndex('M'),
			'N': self.GetColIndex('N'),
			'O': self.GetColIndex('O'),
			'P': self.GetColIndex('P'),
			'Q': self.GetColIndex('Q'),
			'R': self.GetColIndex('R'),
			'S': self.GetColIndex('S'),
			'T': self.GetColIndex('T'),
			'U': self.GetColIndex('U'),
			'V': self.GetColIndex('V'),
			'W': self.GetColIndex('W'),
			'X': self.GetColIndex('X'),
			'Y': self.GetColIndex('Y'),
			'Z': self.GetColIndex('Z')
		}

	def GetColIndex(self, colName):
		colIndex = ord(colName) - ord('A')
		return colIndex

	def GetValue(self, rowIndex, colName):
		colIndex = self.columNameIndexMap[colName]
		return self.sheet.cell_value(rowIndex, colIndex)
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
import numpy as np

def process_sheet(filename):
	
	wb = load_workbook('data/' + filename)
	sheets = wb.sheetnames

	for sheet in sheets:
		ws = wb[sheet]
		#ws.title = ws.title + ' ' + ws['O2'].value

		# read data
		seedlings = []
		for row in range(3,13):
			for col in range(3,13):
				letter = get_column_letter(col)
				seedlings.append(ws[letter + str(row)].value)
		seedlings = np.array(seedlings)

		fillable_names = []
		for row in range(2,6):
			fillable_names.append(ws['N' + str(row)].value)

		fillable_data = []
		for row in range(2,6):
			fillable_data.append(ws['O' + str(row)].value)

		fillable = dict(zip(fillable_names,fillable_data))

		# calculation of metrics
		metrics = {}

		metrics['SIEMBRA'] = fillable['GOLPE']*fillable['TAMAÑO']
		metrics['PROGRENIE'] = np.sum(seedlings)
		metrics['FERTILIDAD'] = metrics['PROGRENIE'] / metrics['SIEMBRA']
		metrics['INVERSION'] =  metrics['SIEMBRA'] / metrics['PROGRENIE']
		metrics['AREA'] = fillable['TAMAÑO'] * fillable['LOSA']
		metrics['DENSIDAD'] = metrics['SIEMBRA'] / metrics['AREA']
		metrics['OCUPACION'] = metrics['AREA'] / metrics['SIEMBRA']
		metrics['COLONIZACION'] = metrics['PROGRENIE'] / fillable['TAMAÑO']
		metrics['RENDIMIENTO'] = metrics['PROGRENIE'] / metrics['AREA']
		metrics['HABITAT'] = metrics['AREA'] / metrics['PROGRENIE']

		# write data into excel sheet
		for row,value in zip(range(2,12),metrics.values()):
			ws['S' + str(row)] = value

	wb.save('data/' + filename)

process_sheet('sample_entry.xlsx')

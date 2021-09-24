from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
import numpy as np
import datetime as dt

fillable_names = ['CULTIVO', 'TAMAÑO', 'LOSA', 'GOLPE']
fillable_units = ['cubos', 'cm2', 'semillas/cubo']
metrics_units = ['semillas', 'plantas', 'plantas/semilla', 'semillas/planta', 'cm2', 'semillas/cm2', 'cm2/semilla', 'plantas/cubo', 'plantas/cm2', 'cm2/planta']
metrics_names = ['SIEMBRA', 'PROGRENIE', 'FERTILIDAD', 'INVERSION', 'AREA', 'DENSIDAD', 'OCUPACION', 'COLONIZACION', 'RENDIMIENTO', 'HABITAT']

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

def write_sheet(parameters):

	wb = Workbook()
	ws = wb.active
	date = dt.datetime.now().strftime('%Y-%m-%d')
	ws.title = date

	ws['B1'] = 'SEMILLERO'

	crop, nrow, ncol = parameters

	ws.merge_cells('B1:' + chr(ord('B')+ncol) + '1')

	for col in range(3,3+ncol):
		ws[get_column_letter(col) + str(2)] = chr(ord('A')+(col-3))

	for row in range(3,3+nrow):
		ws['B' + str(row)] = str(row-2)

	for row in range(2,6):
		ws[get_column_letter(4+ncol) + str(row)] = fillable_names[row-2]

	for row in range(3,6):
		ws[get_column_letter(6+ncol) + str(row)] = fillable_units[row-3]

	for row in range(2,12):
		ws[get_column_letter(8+ncol) + str(row)] = metrics_names[row-3]

	for row in range(2,12):
		ws[get_column_letter(10+ncol) + str(row)] = metrics_units[row-2]

	ws[get_column_letter(5+ncol) + '2'] = crop
	ws[get_column_letter(5+ncol) + '3'] = nrow * ncol

	wb.save(f'data/{crop}_{date}_{nrow}x{ncol}.xlsx')

write_sheet(('cebolla',3,7))
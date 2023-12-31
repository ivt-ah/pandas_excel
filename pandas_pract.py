import pandas as pd 
import xlrd


COLS_TO_PARSE = [0, 2, 6]
COL_TITLES_BEFORE_PIVOT = ['Time Stamp', 'Loop Name', 'Process Value']
STARTROW = 3
NUM_PARAMS = 6

COL_TITLES_AFTER_PIVOT = ['Time Stamp', 'pH', 'DO', 'Agitation', 'CO2', 'O2', 'Temperature']
# SEQUENCE OF UNITS, order important
UNITS = ['Hour', 'Unit', '%', 'RPM', 'SLPM', 'SLPM', '°C']

# col_rename_dict = {
# 	'1-pH_Dev1': 'pH',
# 	'2-DO_Dev1': 'DO',
# 	'Agitation_Dev1': 'Agitation',
# 	'S-Air_Dev1': 'Aeration',
# 	'S-CO2_Dev1': 'CO2',
# 	'S-O2_Dev1': 'Oxygen',
# 	'Temp_Dev1': 'Temperature',
# 	'1-pH_Dev2': 'pH',
# 	'2-DO_Dev2': 'DO',
# 	'Agitation_Dev2': 'Agitation',
# 	'S-Air_Dev2': 'Aeration',
# 	'S-CO2_Dev2': 'CO2',
# 	'S-O2_Dev2': 'Oxygen',
# 	'Temp_Dev2': 'Temperature',
# }

# create DataFrame object from excel file

def get_concat_dataframe(xlrd_book):

	num_sheets = xlrd_book.nsheets 

	dataframes = []
	for sheet in range(num_sheets):
		skiprows = 5 if sheet == 0 else 0
		df = create_dataframe(xlrd_book, sheet, skiprows) 
		dataframes.append(df)

	bioreactor_data = pd.concat(dataframes) \
					    .dropna() \
					    .pivot(index='Time Stamp', columns='Loop Name', values='Process Value')

	# drop any unwanted columns and rename remaining columns
	cols_to_drop = [col for col in list(bioreactor_data.columns) if 'Pump' in col]
	bioreactor_data = bioreactor_data.drop(columns=cols_to_drop)  
									# .rename(columns=col_rename_dict)

	return bioreactor_data



def create_dataframe(xlrd_book, sheet_name, skiprows):
	return pd.read_excel(
		xlrd_book,
		sheet_name=sheet_name,
		skiprows=skiprows,
		usecols=COLS_TO_PARSE,
		header=None,
		names=COL_TITLES_BEFORE_PIVOT,
	)



def create_output_workbook(file_path):

	input_workbook = xlrd.open_workbook(file_path)

	title = input_workbook.sheet_by_index(0).cell_value(rowx=1, colx=1) \
				.split(':')[0]

	# NEED TO CLOSE XLRD WORKBOOK?????

	data = get_concat_dataframe(input_workbook)

	# write data to excel sheet
	writer = pd.ExcelWriter(f'{title}.xlsx', engine='xlsxwriter')
	data.to_excel(writer, 
		          sheet_name='Sheet1', 
		          startrow=STARTROW, 
		          freeze_panes=(STARTROW + 1,NUM_PARAMS),
		          )
	
	# create workbook 
	wb = writer.book

	# add title to cell A1 in Sheet1
	ws = wb.get_worksheet_by_name('Sheet1')
	bold = wb.add_format({'bold': True})
	
	ws.write('A1', title, bold)

	for col, name in enumerate(COL_TITLES_AFTER_PIVOT):
		ws.write(STARTROW - 1, col, name, bold)

	for col, unit in enumerate(UNITS):
		ws.write(STARTROW, col, unit, bold)

	# copy columns headers to row 2 0-indexed
	# write units to row 3 0-indexed
	# units = ['hour', 'Unit', '%', 'RPM', '%', 'SLPM', 'SLPM', '°C']
	# for unit in units:
	# 	ws.write()

	# create chartsheet
	cs = wb.add_chartsheet() 

	# create chart object
	chart = wb.add_chart({ 'type': 'scatter', })

	# configure chart
	chart = configure_chart(chart, data.shape, title)

	# add chart to chartsheet
	cs.set_chart(chart)

	writer.close()


def add_series_to_chart(chart, data_shape):

	# configure the series from the dataframe data
	(max_row, max_col) = data_shape

	for col in range(1, max_col):

		has_y2_axis = col == 3
		series_config = get_series_config(col, max_row, has_y2_axis, STARTROW)

		chart.add_series(series_config)

	return chart


def get_series_config(col, max_row, has_y2_axis, startrow):

	'''
	name = column title
	categories = time stamp values
	values = column values in {name}
	[{sheet_title}, row_min, col_min, row_max, col_max]
	'''
	config = {
			'name': ['Sheet1', startrow, col], # row, col
	 		'categories': ['Sheet1', startrow + 2, 0, startrow + max_row - 1, 0], # max_row - 1 because 0-indexed
	 		'values': ['Sheet1', startrow + 2, col, startrow + max_row - 1, col], 
	 		'marker': { 
	 			'type': 'automatic',
	 			'size': 1, 
	 		 },
	 		 'line': { 
	 		 	'dash_type': 'solid',
	 		 	'size': 1,
	 		 },
	}

	return config | { 'y2_axis': 1 } if has_y2_axis else config


def configure_chart(chart, data_shape, chart_title):

	# configure chart axes and title
	chart.set_x_axis({ 'name': 'EFT, hour', })
	chart.set_y_axis({ 'name': 'DO, %', })
	chart.set_y2_axis({ 'name': 'Agitation, RPM'})

	chart.set_title({ 'name': f'{chart_title} Fermentor Conditions', })
	chart.set_legend({ 'position': 'bottom' })

	# configure the series from the dataframe data
	chart = add_series_to_chart(chart, data_shape)


	return chart





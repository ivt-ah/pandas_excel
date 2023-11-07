import pandas as pd 
# import matplotlib.pyplot as plt
# from xlsxwriter import Workbook

excel_file = 'test.xls'

# B2
# F4_Klebsiella_K102_02agf_23aug2023: DataLog SP, PV, Output, 											

# Read Excel and select a single cell (and make it a header for a column)
title = pd.read_excel(excel_file, 'Sheet1', index_col=None, usecols = "B", header = 10, nrows=0)

serotype = title.columns.values[0].split(':, 1')[0]

COLS_TO_PARSE = [0, 2, 6]
COL_TITLES = ['Time Stamp', 'Loop Name', 'Process Value']

# create DataFrame object from excel file

df_sheet1 = pd.read_excel(
	excel_file, 
	sheet_name=0, 
	skiprows=5,
	usecols=COLS_TO_PARSE,
	header=None,
	names=COL_TITLES,
	)
df_sheet2 = pd.read_excel(
	excel_file, 
	sheet_name=1, 
	usecols=COLS_TO_PARSE,
	header=None,
	names=COL_TITLES,
	)
bioreactor_data = pd.concat([df_sheet1, df_sheet2]) \
					.dropna() \
					.pivot(index='Time Stamp', columns='Loop Name', values='Process Value')


cols_to_drop = [col for col in list(bioreactor_data.columns) if 'Pump' in col]
bioreactor_data = bioreactor_data.drop(columns=cols_to_drop)

col_rename_dict = {
	'1-pH_Dev1': 'pH',
	'2-DO_Dev1': 'DO',
	'Agitation_Dev1': 'Agitation',
	'S-Air_Dev1': 'Aeration',
	'S-CO2_Dev1': 'CO2',
	'S-O2_Dev1': 'Oxygen',
	'Temp_Dev1': 'Temperature',
}

bioreactor_data = bioreactor_data.rename(columns=col_rename_dict)


# write data to excel sheet

writer = pd.ExcelWriter('test_graphs2.xlsx', engine='xlsxwriter')
bioreactor_data.to_excel(writer, sheet_name='Sheet1')
wb = writer.book
ws = writer.sheets['Sheet1']

# create a chart object
chart = wb.add_chart({ 'type': 'scatter', })

# configure chart axes and title
chart.set_x_axis({ 'name': 'EFT, min', })
chart.set_y_axis({ 'name': 'Agitation (RPM), DO (%)', })
chart.set_title({ 'name': f'{serotype} Fermentor Conditions', })

# configure the series from the dataframe data
(max_row, max_col) = bioreactor_data.shape

for col in range(1, max_col):
	chart.add_series({
		'name': ['Sheet1', 0, col],
 		'categories': ['Sheet1', 1, 0, max_row, 0],
 		'values': ['Sheet1', col, max_col, max_row, col],
 		'marker': { 
 			'type': 'automatic',
 			'size': 1, 
 		 },
 		 'line': { 
 		 	'dash_type': 'solid',
 		 	'size': 1,
 		 },
	})

ws.insert_chart('C1', chart)

writer.close()


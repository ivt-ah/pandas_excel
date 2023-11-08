import pandas as pd 
import os
import glob
import xlrd
from pandas_pract import create_output_workbook, get_concat_dataframe

'''

# use glob to get all excel files
dir_name = 'excel_files'
path = f'{os.getcwd()}\\{dir_name}' 

excel_files = glob.glob(os.path.join(path, '*.xls'))


# loop over list of excel files
for excel_file in excel_files:

	input_workbook = xlrd.open_workbook(excel_file)

	# get title from cell 'B2'
	title = input_workbook.sheet_by_index(0).cell_value(rowx=1, colx=1) \
				.split(':')[0]
	
	create_output_workbook(title, input_workbook)

'''
dir_name = 'excel_files'
path = f'{os.getcwd()}\\{dir_name}'
file_path = os.path.join(path, 'a.xls')

# TEST SINGLE FILE
# input_workbook = xlrd.open_workbook(file_path)
# df = get_concat_dataframe(input_workbook)
# print(df.shape)
# print(df.head())
	
create_output_workbook(file_path)


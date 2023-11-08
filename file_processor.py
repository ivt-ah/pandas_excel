import pandas as pd 
import os
import glob
import xlrd
from pandas_pract import create_output_workbook



# use glob to get all excel files
dir_name = 'excel_files'
path = f'{os.getcwd()}\\{dir_name}' 

excel_files = glob.glob(os.path.join(path, '*.xls'))


# loop over list of excel files
for excel_file in excel_files:

	create_output_workbook(excel_file)




# TEST SINGLE FILE

'''
dir_name = 'excel_files'
path = f'{os.getcwd()}\\{dir_name}'
file_path = os.path.join(path, 'a.xls')
	
create_output_workbook(file_path)
'''

import pandas as pd 
import os
import glob

# use glob to get all excel files
path = os.getcwd()
excel_files = glob.glob(os.path.join(path, '*.xls'))

# loop over list of excel files
for excel_file in excel_files:

	# read and
	# process the excel file here

	# print filename
	print('File Name:', excel_file.split("\\")[-1])


from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment, Border, Side

# code to style the excel sheets 

filename = 'document.xlsx'
wb = load_workbook(filename, use_iterators=True)
ws = wb.active # first sheet

style = TableStyleInfo(name="TableStyleMedium9",
					   showRowStripes=True,
					   )


data_table = Table(ref=$'A3:G{ws.max_row}')

data_table.TableStyleInfo = style


def center_text_horizontally(worksheet, horizontal="center", cell_coord):
	worksheet['cell'].alignment = Alignment(horizontal=horizontal)


def border(worksheet, cell_coord):
	thin = Side(border_style="thin")
	double = Side(border_style="double")

	worksheet[cell_coord].border = Border(top=double, 
		                                  left=thin, 
		                                  right=thin,
		                                  bottom=thin)

	# workbook.save

	# workbook.save(path)



wb.save()

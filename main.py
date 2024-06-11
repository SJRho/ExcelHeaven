import xlwings as xw, pandas as pd, numpy as np
import os, shutil, time, random, winsound

def current_selected_cell(wb):
	return wb.app.selection.get_address(row_absolute=False,column_absolute=False)



def main():
	# Create obstacle course .xlsx workbook
	# dstfilepath = 'ExcelHeaven.xlsx'
	#shutil.copyfile(srcfilepath, dstfilepath)
	# wb = ...
	# Open workbook
	#wb = xw.Book(dstfilepath)
      
	wb = xw.Book('1_COPY.xlsx')
	s = current_selected_cell(wb)
	# Game Loop
	while len(xw.apps) > 0:
		if s != current_selected_cell(wb):
			winsound.Beep(450,350)
			s = current_selected_cell(wb)




if __name__ == "__main__":
    main()
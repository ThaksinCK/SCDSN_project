import openpyxl 
import datetime
from openpyxl import Workbook

x =datetime.datetime.now()
autows = x.strftime("%B")
wb = openpyxl.load_workbook('')
ws = wb[str(autows)]

def scanrow(room):
	countrow = 0
	for row in ws.iter_rows(1,10,1,1):
		for cell in row:
			countrow += 1
			if cell.value == room:
				return countrow

def scancol():
  countcol = 97
  time = datetime.datetime.now()
	timecal = str(time.year) + '-' + str(time.month) + '-' + str(time.day)
	
	for row in ws.iter_rows(1,1,1,15):
		for cell in row:
			x = str(cell.value)
			y = x.replace(' 00:00:00','')
			countcol += 1
			if y == timecal:
				return chr(countcol)

def scanroom(room):				
	row = scanrow(room)
	col = scancol()
	cell = str(col)+str(row)
	return cell

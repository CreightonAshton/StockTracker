from bs4 import BeautifulSoup
from urllib.request import urlopen
from datetime import datetime
from openpyxl import Workbook
from openpyxl import load_workbook

current_date = "{:%B %d, %Y}".format(datetime.now())
current_month = "{:%B}".format(datetime.now())
num_month = "{:%m}".format(datetime.now())
current_year = "{:%Y}".format(datetime.now())
#my_portfolio = {}

def update_stocks (wb_name):
	wb = load_workbook(wb_name)
	ws = wb.active
	num_stocks = 10
	for i in range(2,num_stocks+2):
		cell_of_symbol = 'A' + str(i)
		my_stock_name = ws[cell_of_symbol].value
		content = urlopen('http://finance.yahoo.com/q?s=' + my_stock_name)
		page_holder = content.read()
		content.close()
		soup = BeautifulSoup(page_holder)
		span_id = 'yfs_l84_' + my_stock_name.lower()
		my_stock_price = float(soup.find(id=span_id).string)
		print ("Current value of ", my_stock_name, " is $", my_stock_price, ".", sep="")
		#my_portfolio[my_stock_name] = my_stock_price
		cell_of_current_price = 'S' + str(i)
		ws[cell_of_current_price] = my_stock_price
	ans = ""
	while ans != 'y' or ans != 'n':
		ans = str(input("Would you like to save your portfolio? (y/n) "))
		if ans == 'y':
			wb.save(num_month + ' - ' + current_month + '/My Index ' + current_date + '.xlsx')
			print("Portfolio has been saved.")
			break
		elif ans == 'n':
			break
		else:
			print("Please enter either y or n")
"""
def save_stocks(wb_name):
	wb = load_workbook(wb_name)
	wb.save(num_month + ' - ' + current_month + '/My Index ' + current_date + '.xlsx')
	print("Portfolio has been saved.")
"""
def new_portfolio():
	user_name = str(input("What is your name? "))
	num_stocks = int(input("How many stocks do you want to track? "))
	wb = Workbook()
	ws = wb.active
	ws.title = user_name + "'s Portfolio"
	print(ws.title)
	
	
def startup():	
	new_old = ""
	while new_old != 'new' or new_old != 'open':
		new_old = str(input("New file: new    Open existing file: open \n"))
		if new_old == "new":
			new_portfolio()
			break
		elif new_old == "open":
			wb_name = str(input("What is the name of your portfolio? "))
			file_path = wb_name + '.xlsx'
			update_stocks(file_path)
			break
		else:
			print ("Please enter either new or open.")

startup()			
#update_stocks('My Index ' + current_year + '.xlsx')
#save_stocks('My Index ' + current_year + '.xlsx')


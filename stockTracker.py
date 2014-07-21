from bs4 import BeautifulSoup
from urllib.request import urlopen
from datetime import datetime

num_stocks = int(input("How many stocks do you want to track? "))
my_portfolio = {}

for i in range(num_stocks):
	my_stock_name = str(input("Name of stock: ")).upper()
	content = urlopen('http://finance.yahoo.com/q?s=' + my_stock_name)
	page_holder = content.read()
	content.close()
	soup = BeautifulSoup(page_holder)
	span_id = 'yfs_l84_' + my_stock_name.lower()
	span_list = soup.find_all('span')
	my_stock_price = float(soup.find(id=span_id).string)
	print ("Current value of ", my_stock_name, " is $", my_stock_price, ".", sep="")
	my_portfolio[my_stock_name] = my_stock_price

current_date = "{:%B %d, %Y}".format(datetime.now())
folio = open("portfolios\my_portfolio.txt", "a")
folio.write(str(current_date) + "\n")
for stock, price in my_portfolio.items():
	string_to_write = str(stock) + ": " + str(price) + "\n"
	folio.write(string_to_write)
folio.write("\n")
folio.close()
print("Portfolio has been saved")
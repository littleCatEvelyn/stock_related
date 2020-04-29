import argparse
from datetime import date, timedelta
import pandas_datareader.data as pdr
# import time
import openpyxl

stock_list1 = ["AAPL", "MSFT", "FB", "BABA", "DAL", "", "BAM", "MAR", "BILI", "SE", "EOG", "", "IVV", "PDD", ""]
time = date.today()
while time.weekday() > 4: # Mon-Fri are 0-4
	time -= timedelta(days=1)
COL_INDEX = 12
ROW_INDEX1 = 4
path = None

def crawl_prices(stock_list):
	list_len = len(stock_list)
	prices = []
	for i in range(list_len):
		if (stock_list[i] != ""):
			value = pdr.get_data_yahoo(stock_list[i], time, time)
			price = list(value["Close"].values)
			prices.append(price[0])
		else:
			prices.append("")
	return prices

def fill_prices(prices, row_start):
	wb = openpyxl.load_workbook(filename=path)
	ws = wb['Overview']
	for row, entry in enumerate(prices, start=row_start):
		if (entry == ""):
			continue
		ws.cell(row=row, column=COL_INDEX, value=round(entry,2))
	wb.save(path)

def crawl_and_fill():
	prices1 = crawl_prices(stock_list1)
	fill_prices(prices1, ROW_INDEX1)

def main():
	crawl_and_fill()


if __name__ == '__main__':
	parser = argparse.ArgumentParser()
	parser.add_argument('--path', '-p', type=str, default='../plan.xlsx')
	args = parser.parse_args()
	path = args.path

	main()
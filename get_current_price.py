import argparse
import datetime
import pandas_datareader.data as pdr
# import time
import openpyxl

stock_list1 = ["IVV", "AAPL", "MSFT", "FB", "BABA"]
stock_list2 = ["MAR", "BILI", "SE"]
time = datetime.date.today()
COL_INDEX = 17
ROW_INDEX1 = 3
ROW_INDEX2 = 10
path = None

def crawl_prices(stock_list):
	list_len = len(stock_list)
	prices = []
	for i in range(list_len):
		value = pdr.get_data_yahoo(stock_list[i], time, time)
		price = list(value["Close"].values)
		prices += price
	return prices

def fill_prices(prices, row_start):
	wb = openpyxl.load_workbook(filename=path)
	ws = wb.get_sheet_by_name('Sheet1')
	for row, entry in enumerate(prices, start=row_start):
		ws.cell(row=row, column=COL_INDEX, value=round(entry,2))
	wb.save(path)

def crawl_and_fill():
	prices1 = crawl_prices(stock_list1)
	prices2 = crawl_prices(stock_list2)
	fill_prices(prices1, ROW_INDEX1)
	fill_prices(prices2, ROW_INDEX2)

def main():
	# while(True):
	# 	crawl_and_fill()
	# 	time.sleep(60)
	crawl_and_fill()


if __name__ == '__main__':
	parser = argparse.ArgumentParser()
	parser.add_argument('--path', '-p', type=str, default='./plan.xlsx')
	args = parser.parse_args()
	path = args.path

	main()
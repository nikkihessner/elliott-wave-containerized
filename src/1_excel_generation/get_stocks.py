#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os 
import subprocess
from openpyxl import Workbook
from openpyxl.worksheet.formula import ArrayFormula
from datetime import date, timedelta
import time
import pandas as pd
import xlwings as xw

data = pd.read_excel('data/exchanges/industry.xlsx')
print(data.head())

df = pd.DataFrame(data[0:])
print(df.head())
print(df.count())
print(len(df))

class LoadStocks():
    def nasdaq():
        nasdaq = pd.read_excel('data/exchanges/nasdaq.xlsx')
        return nasdaq

    def nyse():
        nyse = pd.read_excel('data/exchanges/nyse.xlsx')
        return nyse

    def etfs():
        etf = pd.read_excel('data/exchanges/etf.xlsx')
        return etf

    def create_filepath(stock):
        output_folder = 'data/stocks'
        file_name = f'{stock}.xlsx'
        excel_file_path = os.path.join(output_folder, file_name)
        return excel_file_path

class StockHistory:
    def create_stockhistory(stock, excel_file_path, nasdaq, nyse, etf):
        if stock == "BRK-B":
            stock_symbol = f"XNYS:BRK.B"
        elif stock in nasdaq.iloc[:, 1].values:
            stock_symbol = f"XNAS:{stock}"
        elif stock in nyse.iloc[:, 1].values:
            stock_symbol = f"XNYS:{stock}"
        elif stock in etf.iloc[:, 1].values:
            stock_symbol = f"ARCX:{stock}"
        else:
            stock_symbol = None

        if stock_symbol is not None:
            # Open the Excel file
            wb = xw.Book()
            
            span = 365 * 10
            start_date = "TODAY()" + "-" + str(span)
            end_date = "TODAY()"

            # Access the specific worksheet
            sheet = wb.sheets['Sheet1']  # Replace 'Sheet1' with your sheet name

            # Specify the cell for the STOCKHISTORY function
            cell_range = sheet.range('A1' + ':' + 'F3650')
            stockhistory_formula = f'=STOCKHISTORY("{stock_symbol}", {start_date}, {end_date}, 0, 1, 0, 1, 2, 3, 4, 5)'  # Example STOCKHISTORY formula

            # Input the STOCKHISTORY formula as an array formula to prevent the '@' symbol
            cell_range.formula_array = stockhistory_formula
            wb.app.calculate()
            time.sleep(3)

            # Save and close the workbook
            wb.save(excel_file_path)
            wb.close()
            
        return stock_symbol

    def generate_data():
        count = 0
        nasdaq = LoadStocks.nasdaq()
        nyse = LoadStocks.nyse()
        etf = LoadStocks.etfs()
        app = xw.App(visible=False)

        for stock in df.iloc[:len(df), 1]:
            try:
                excel_file_path = LoadStocks.create_filepath(stock)
                stock_symbol = StockHistory.create_stockhistory(stock, excel_file_path, nasdaq, nyse, etf)
                if stock_symbol is not None:
                    count += 1

                if count == 100:
                    app.quit()
                    ExcelFile.close()
                    time.sleep(3)
                    app = xw.App(visible=False)
                    count = 0
            except:
                continue
        return app

class ExcelFile():
    def __init__(self, excel_file_path):
        self.process = subprocess.Popen(['open', excel_file_path])
        self.pid = self.process.pid
    
    @staticmethod
    def close():
        applescript = 'tell application "Microsoft Excel" to quit saving no'
        subprocess.run(['osascript', '-e', applescript])

def main():
    app = StockHistory.generate_data()
    app.quit()
    folder_path = 'data/stocks'
    count = 0
    for filename in os.listdir(folder_path):
        try: 
            file_path = os.path.join(folder_path, filename)
            print(file_path)
            data = pd.read_excel(file_path, header=0)
            df = pd.DataFrame(data)
            if df.shape[1] != 6:
                StockHistory.generate_data()
            else:
                continue
        except ValueError:
            print(filename)
            count += 1

    print("Files with Errors: " + str(count))
    return 

if __name__ == "__main__":
    main()

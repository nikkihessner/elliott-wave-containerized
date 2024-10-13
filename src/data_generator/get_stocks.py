#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os 
import subprocess
import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.formula import ArrayFormula
import datetime
from datetime import date, timedelta
import time
import pandas as pd
import pyautogui
# import keyboard
import signal
import xlwings as xw


# In[2]:


data = pd.read_excel('../source/industry.xlsx')
data.head()


# In[3]:


df = pd.DataFrame(data[0:])
df.head()
df.count()
len(df)


# In[4]:


def clear_old_data(folder_path):
    today = date.today()
    for filename in os.listdir('../data'):
        file_path = os.path.join(folder_path, filename)
        try:
            creation_time = datetime.date.fromtimestamp(os.path.getctime(file_path))

            if creation_time < today:
                os.remove(file_path)
                print(f"Deleted {filename}")
        except Exception as e:
            print(f"Error deleting {filename}: {e}")


# In[5]:


def load_nasdaq():
    print("Nasdaq")
    nasdaq = pd.read_excel('../source/nasdaq.xlsx')
    print(nasdaq.iloc[:, 1])
    return nasdaq


# In[6]:


def load_nyse():
    nyse = pd.read_excel('../source/nyse.xlsx')
    print(nyse.iloc[:, 1])
    return nyse


# In[7]:


def load_etf():
    etf = pd.read_excel('../source/etf.xlsx')
    return etf


# In[8]:


def create_filepath(stock):
    output_folder = '../data/'
    # csv_folder = '../csv/'
    file_name = f'{stock}.xlsx'
    # csv_name = f'{stock}.xlsx'
    excel_file_path = os.path.join(output_folder, file_name)
    # csv_file_path = os.path.join(csv_folder, csv_name)
    return excel_file_path


# In[9]:


# def create_stockhistory(stock, excel_file_path, nasdaq, nyse, etf, app):
#     # Open the Excel file
#     # Set to True if you want to see the Excel application
#     wb = xw.Book()

#     if stock == "BRK-B":
#         stock_symbol = f"XNYS:BRK.B"
#     elif stock in nasdaq.iloc[:, 1].values:
#         stock_symbol = f"XNAS:{stock}"
#     elif stock in nyse.iloc[:, 1].values:
#         stock_symbol = f"XNYS:{stock}"
#     elif stock in etf.iloc[:, 1].values:
#         stock_symbol = f"ARCX:{stock}"
#     else:
#         stock_symbol = None
    
#     if stock_symbol is not None:
#         span = 365 * 10
#         start_date = date.today() - timedelta(days=span)
#         end_date = date.today()

#         # Access the specific worksheet
#         sheet = wb.sheets['Sheet1']  # Replace 'Sheet1' with your sheet name

#         # Specify the cell for the STOCKHISTORY function
#         cell_range = sheet.range('A1' + ':' + 'F3650')
#         stockhistory_formula = f'=STOCKHISTORY("{stock_symbol}", "{start_date}", "{end_date}", 0, 1, 0, 1, 2, 3, 4, 5)'  # Example STOCKHISTORY formula

#         # Input the STOCKHISTORY formula as an array formula to prevent the '@' symbol
#         cell_range.formula_array = stockhistory_formula
#         wb.app.calculate()
#         time.sleep(3)

#         # # Load data from Excel into a pandas DataFrame
#         # data_range = sheet.range('A1:F3650')
#         # data = data_range.options(pd.DataFrame, header=True, index=False).value

#         # # Save DataFrame to .csv file
#         # data.to_csv(csv_file_path, index=False)

#         # Save and close the workbook
#         wb.save(excel_file_path)
#         time.sleep(1)

#         try:
#             wb.close()
#             time.sleep(1)
#         except Exception as e:
#             app.kill()
#             return
#         return stock_symbol
#     else:
#         app.kill()
#         time.sleep(3)
#         return stock_symbol


# In[10]:


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

        # # Load data from Excel into a pandas DataFrame
        # data_range = sheet.range('A1:B300')
        # data = data_range.options(pd.DataFrame, header=True, index=False).value

        # # Save DataFrame to .csv file
        # data.to_csv(csv_file, index=False)

        # Save and close the workbook
        wb.save(excel_file_path)
        # time.sleep(1)
        wb.close()
        # time.sleep(1)
        # time.sleep(3)
        
    return stock_symbol


# In[11]:


# stock = "BRK-B"
# excel_file_path = create_filepath(stock)
# create_stockhistory(stock, excel_file_path)
# convert_values(excel_file_path)
# def pause_resume_loop(function):
#     paused = False 
#     while True: 
#         if paused:
#             print("Loop paused...")
#         else:
#             function()
#         if keyboard.is_pressed(' '):
#             paused = not paused
#             if paused:
#                 print("Loop paused...")
#             else:
#                 print("Loop resumed...")
#             while keyboard.is_pressed(' '):
#                 time.sleep(0.1)
        
#         time.sleep(0.1)




# In[12]:


def excel_function_bar_click():
    # Move the mouse to the desired location and click
    pyautogui.click(x=59, y=222)
    pyautogui.click(x=246, y=174)
    # Press the Enter key
    pyautogui.press('enter')
    time.sleep(1.5)

def select_all_click():
    # Select All
    pyautogui.click(x=65, y=250)
    time.sleep(0.2)
    pyautogui.click(x=164, y=9)
    time.sleep(0.2)
    pyautogui.click(x=175, y=253)
    time.sleep(1.5) 

def copy_click():
    # Copy
    pyautogui.click(x=65, y=250, button='right')
    time.sleep(0.2)
    clicks = 0
    while clicks < 2: 
        pyautogui.keyDown('down')
        pyautogui.keyUp('down')
        clicks += 1
    pyautogui.press('enter')   

def paste_special():
    # Paste Special, Formulas
    pyautogui.click(x=51, y=250)
    time.sleep(0.2)
    pyautogui.click(x=51, y=250)
    time.sleep(0.2)
    pyautogui.click(x=51, y=250, button='right')

    clicks = 0 
    while clicks < 4:
        pyautogui.keyDown('down')
        pyautogui.keyUp('down')
        clicks += 1

    pyautogui.keyDown('right')
    pyautogui.keyUp('right')

    clicks = 0
    while clicks < 7:
        pyautogui.keyDown('down')
        pyautogui.keyUp('down')
        clicks += 1 

    pyautogui.press('enter')
    time.sleep(1)

def excel_save():
    # Save
    pyautogui.click(x=122, y=12)
    clicks = 0
    while clicks < 6:
        pyautogui.keyDown('down')
        pyautogui.keyUp('down')
        clicks += 1 
    pyautogui.press('enter')
    time.sleep(1)

def excel_close():
    # Close
    pyautogui.click(x=21, y=40)
    pyautogui.click(x=21, y=40)
    # clicks = 0
    # while clicks < 5:
    #     pyautogui.keyDown('down')
    #     pyautogui.keyUp('down')
    #     clicks += 1 
    # pyautogui.press('enter')   


# In[13]:


class ExcelFile():
    def __init__(self, excel_file_path):
        self.process = subprocess.Popen(['open', excel_file_path])
        self.pid = self.process.pid
    
    @staticmethod
    def close():
        applescript = 'tell application "Microsoft Excel" to quit saving no'
        subprocess.run(['osascript', '-e', applescript])



# In[14]:


def convert_values(excel_file_path, t):
    # Check if the file exists
    if os.path.exists(excel_file_path):
        # Open the Excel file using the default application
        file = ExcelFile(excel_file_path) 
        # subprocess.Popen(['open', excel_file_path])
    else:
        print(f"File '{excel_file_path}' does not exist.")
    # Wait for a few seconds to give time to switch to the desired application
    time.sleep(t)

    pyautogui.click(x=1103, y=537)
    time.sleep(0.2)
    pyautogui.click(x=727, y=394)
    time.sleep(0.2)
    pyautogui.click(x=158, y=850)
    time.sleep(1)

    excel_function_bar_click()

    select_all_click()

    copy_click()

    paste_special()

    pyautogui.click(x=705, y=405)

    excel_save()

    excel_close()

    return


# In[15]:


def generate_data():
    count = 0
    nasdaq = load_nasdaq()
    nyse = load_nyse()
    etf = load_etf()
    app = xw.App(visible=False)

    for stock in df.iloc[:len(df), 1]:
        try:
            excel_file_path = create_filepath(stock)
        # if os.path.exists(excel_file_path):
        #     print(stock)
        #     count += 1
        # #     print(count)
        # else:
            stock_symbol = create_stockhistory(stock, excel_file_path, nasdaq, nyse, etf)
            if stock_symbol is not None:
            # #     # convert_values(excel_file_path, t)
                count += 1
            # # #     # time.sleep(2)

            if count == 100:
                app.quit()
                ExcelFile.close()
                time.sleep(3)
                app = xw.App(visible=False)
                count = 0
        except:
            continue
        # t = delay
        # if stock_symbol is not None:
        # # #     # convert_values(excel_file_path, t)
        #     count += 1
        # # # #     # time.sleep(2)

        # if count == 100:
        #     ExcelFile.close()
        # # #     pyautogui.click(x=158, y=850)
        # # #     time.sleep(3)
        # #     count = 0


# In[16]:


def main():
    generate_data()
    app.quit()
    folder_path = '../data/'
    count = 0
    for filename in os.listdir(folder_path):
        try: 
            file_path = os.path.join(folder_path, filename)
            print(file_path)
            data = pd.read_excel(file_path, header=0)
            df = pd.DataFrame(data)
            if df.shape[1] != 6:
                generate_data()
            else:
                continue
        except ValueError:
            print(filename)
            count += 1

    print("Files with Errors: " + str(count))
    return 


# In[17]:


folder_path = '../data/'
clear_old_data(folder_path)


# In[18]:


if __name__ == "__main__":
    main()


# In[ ]:





# In[ ]:


# tracking = True 
# while tracking:
#     # Get the current mouse position
#     x, y = pyautogui.position()
#     #236, 191
#     # Print the mouse position
#     time.sleep(5)
#     print(f"Mouse position: x={x}, y={y}")


# In[ ]:


# ExcelFile.close()


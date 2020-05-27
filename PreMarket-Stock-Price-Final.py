import gspread
from oauth2client.service_account import ServiceAccountCredentials
from os import path
from bs4 import BeautifulSoup as Soup
import requests
import PySimpleGUI as sg


DATA_DIR = 'C:/Users/xbsqu/Desktop/Python Learning/Projects/Premarket Stock Price'

#Connecting to G Sheet...

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(path.join(DATA_DIR, 'client_secret.json'), scope)
client = gspread.authorize(creds)

sheet = client.open('Stock Watcher')
worksheet = sheet.get_worksheet(0)

#Now connected to the G Sheet.

row_limit = 1+len(worksheet.col_values(1)) 


for i in range(2, row_limit):
    sg.one_line_progress_meter('Pre-Market $ Lookup', i+1, row_limit, 'key','Fetching premarket prices from Yahoo Finance...')
    
    if worksheet.acell(f'A{i}') == "":
    
            pass
    
    else:
        stock_symbol = worksheet.acell(f'A{i}').value
    
    #Here is where we use a web scraper to grab the premarket price information
    
        url_slug = 'https://finance.yahoo.com/quote/'
        stock_url = url_slug + stock_symbol
      
    try:
        stock_response = requests.get(stock_url)
        stock_soup = Soup(stock_response.text)
    
        stock_pm_price = stock_soup.find('p',{'class': 'fullview-news-outer'}).find('span').text
    
    #Webscraper complete; now we need to write the scraped data to the worksheet and repeat for the remaining rows
        worksheet.update_cell(i, 9, stock_pm_price)
        
    except:      
       worksheet.update_cell(i, 9, "")
        
        

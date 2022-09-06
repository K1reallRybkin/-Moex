import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import xlsxwriter

option = Options()
option.add_argument("--disable-infobars")
browser = webdriver.Chrome('C:\Program Files\Google\chromedriver.exe',chrome_options=option)
browser.get('https://www.moex.com/ru/derivatives/currency-rate.aspx?currency=USD_RUB')

agree = browser.find_element(By.CLASS_NAME,  'btn2-primary')
agree.click()

USD_RUB = Select(browser.find_element(By.ID, 'ctl00_PageContent_CurrencySelect'))
# select by value
USD_RUB.select_by_value('CAD_RUB')

day_USDRUB_1 = Select(browser.find_element(By.ID, 'd1day'))
# select by value
day_USDRUB_1.select_by_value('1')

month_USDRUB_1 = browser.find_element(By.ID, 'd1month')
month_USDRUB_1.click()

year_USDRUB_1 = browser.find_element(By.ID, 'd1year')
year_USDRUB_1.click()

day_USDRUB_2 = browser.find_element(By.ID, 'd2day')
day_USDRUB_2.click()

month_USDRUB_2 = browser.find_element(By.ID, 'd2month')
month_USDRUB_2.click()

year_USDRUB_2 = browser.find_element(By.ID, 'd2year')
year_USDRUB_2.click()

show_USDRUB_1 = browser.find_element(By.CLASS_NAME,  'button80')
show_USDRUB_1.click()


headers = {
    'user-agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.174 '
                  'YaBrowser/22.1.3.942 Yowser/2.5 Safari/537.36',
    'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,'
              'application/signed-exchange;v=b3;q=0.9 '
}

tables = pd.read_html('https://www.moex.com/ru/derivatives/currency-rate.aspx?currency=USD_RUB')
print(tables[2])

result = pd.DataFrame(tables[2]['Дата']['Дата'].values, columns=['Дата'])
result['Значение'] = tables[2]['Курс основного клиринга']['Значение'].values
result['Время'] = tables[2]['Курс основного клиринга']['Время'].values

result['Дата'] = pd.to_datetime(result['Дата'])
result['Время'] = result['Время']
result['Значение'] = result['Значение'].astype(float)

writer = pd.ExcelWriter('result_cur.xlsx', date_format="YYYY-MM-DD", datetime_format="YYYY-MM-DD")
result.to_excel(writer, sheet_name='USD_RUB', index=False, na_rep='NaN')

format_ = writer.book.add_format({'num_format': '#,##0.00 ₽'})
writer.sheets['USD_RUB'].set_column(first_col=1, last_col=2, cell_format=format_)

# Dynamically adjust all the column lengths
for column in result:
    column_length = max(result[column].astype(str).map(len).max(), len(column))
    col_idx = result.columns.get_loc(column)
    writer.sheets['USD_RUB'].set_column(col_idx, col_idx, column_length)

writer.save()
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
#driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

# Read Excel File
df = pd.read_excel('supermarket_sales.xlsx')

# Select columns: 'Gender', 'Product line', 'Total'
df = df[['Gender', 'Product line', 'Total']]
#
# # Make pivot table
pivot_table = df.pivot_table(index='Gender', columns='Product line',
                             values='Total', aggfunc='sum').round(0)

# # Export pivot table to Excel file
pivot_table.to_excel('pivot_table.xlsx', 'Report', startrow=4)

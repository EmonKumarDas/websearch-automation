from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
import datetime
from selenium.webdriver.common.by import By
# Load the Excel file
df = pd.read_excel('E:\web automation assingment\Excel.xlsx', sheet_name=None)

# Get today's day of the week
day_of_week = datetime.datetime.today().strftime('%A')
# Get the data for today's day of the week
data = df[day_of_week]
print(data.columns)
# Initialize the webdriver (make sure to specify the correct path)
driver = webdriver.Chrome()

for index, row in data.iterrows():
    # Go to www.google.com
    driver.get('http://www.google.com')
    time.sleep(2)  # wait for the page to load
    # Find the search box and enter the keyword
    search_box = driver.find_element(By.NAME, 'q')
    search_box.send_keys(row['search'])
    time.sleep(2)  # wait for the options to load

    # Get all options
    options = driver.find_elements(By.CSS_SELECTOR, '.erkvQe li')

    # Find the longest and shortest options
    longest_option = max(options, key=lambda option: len(option.text))
    shortest_option = min(options, key=lambda option: len(option.text))

    # Update the DataFrame
    data.loc[index, 'Longest Option'] = longest_option.text
    data.loc[index, 'Shortest Option'] = shortest_option.text

# Save the updated DataFrame back to the Excel file
with pd.ExcelWriter('E:\web automation assingment\Excel.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    data.to_excel(writer, sheet_name=day_of_week)

driver.quit()

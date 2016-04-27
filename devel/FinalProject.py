#! python3


from openpyxl import Workbook
from selenium import webdriver
import time, os, openpyxl


# Getting the downloaded files from the web....
print('Openning browser.........')
# This is setting the changing the settings of the browser and the path of the downloaded file..
profile = webdriver.FirefoxProfile()
profile.set_preference("browser.download.folderList", 2)
profile.set_preference("browser.download.manager.showWhenStarting", False)
profile.set_preference("browser.download.dir",os.getcwd())
profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
# Downloading the file from the web...
driver = webdriver.Firefox(firefox_profile=profile)
driver.get("https://drive.google.com/open?id=0B37e5yn5lQUQNDJoLXJEYTQ4bkk")
driver.find_element_by_css_selector(".drive-viewer-download-icon").click()
# Giving the computer enough time to download then close the broswer....
# what i have noticed when running my code on a windows platform
# is that i had to adjust the sleeptime giving the computer enought time to dowmload the file
# you can adjust time.sleep(6) on a slower computer 
time.sleep(4)
driver.quit()

# Opening the Downloaded sheet and grabbing the information...
print('Cleaning up the downloaded file ........')
wb = openpyxl.load_workbook('sprint1.xlsx')
sheet = wb.get_sheet_by_name('Sheet1')
wb1 = Workbook()
wbsheet = wb1.active
# Looping through the columns to grab the information..
for cellObj in sheet.columns[0]:
	wbsheet['A' + str(cellObj.row)] = cellObj.value
for cellObj in sheet.columns[1]:
	wbsheet['B' + str(cellObj.row)] = cellObj.value
for cellObj in sheet.columns[14]:
	wbsheet['C' + str(cellObj.row)] = cellObj.value
#Creating a new spreadSheet with the final results..	
wb1.save('Results.xlsx')
print('Done..........')

# import
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
import time

# folder, from/to (format sensitive)
dwd_folder = r'C:\Users\Shithi.Maitra\OneDrive - Unilever\Pictures\Screenshots'
report_from = '29 December 2022'
report_to = '1 January 2023'

# URL
options = webdriver.ChromeOptions()
prefs = {'download.default_directory': dwd_folder}
options.add_experimental_option('prefs', prefs)
driver = webdriver.Chrome(options=options)
driver.get("http://ctgsapp10006/ibp/LoginPage.aspx?LogoutFrom=LoggedOut")
driver.maximize_window()
time.sleep(2)

# login
driver.find_element(By.XPATH, '//*[@id="txtUserID"]').send_keys('000722437')
driver.find_element(By.XPATH, '//*[@id="txtPassword"]').send_keys('unilever123')
driver.find_element(By.XPATH, '//*[@id="btnLogin"]').click()

# claim > damage > material status report
driver.find_element(By.LINK_TEXT, 'Claim').click()
driver.find_element(By.LINK_TEXT, 'Damage').click()
driver.find_element(By.LINK_TEXT, 'Material Status Report').click()

# iframe
driver.switch_to.frame(driver.find_element(By.XPATH, '//*[@id="ifrmPage"]'))

# get date element
def get_date_element(dt, from_to):
    for r in range(3, 9):
        for c in range(1, 8):
            xpath = '//*[@id="ctl00_cphPage_cebc_Date_Range' + from_to + 'Date__Calendar_RootTable"]/tbody/tr/td/table/tbody/tr[' + str(r) + ']/td[' + str(c) + ']'
            elem = driver.find_element(By.XPATH, xpath)
            if elem.text == dt and 'previous' not in elem.get_attribute("class") and 'next' not in elem.get_attribute("class"):
                return elem

# from
report_from = report_from.split()
driver.find_element(By.XPATH, '//*[@id="ctl00_cphPage_cebc_Date_RangeFromDate_down"]').click()
dropdown = driver.find_element(By.XPATH, '//*[@id="ctl00_cphPage_cebc_Date_RangeFromDate__Calendar_YearsDropDown"]')
dd = Select(dropdown)
dd.select_by_visible_text(report_from[2])
dropdown = driver.find_element(By.XPATH, '//*[@id="ctl00_cphPage_cebc_Date_RangeFromDate__Calendar_MonthsDropDown"]')
dd = Select(dropdown)
dd.select_by_visible_text(report_from[1])
get_date_element(report_from[0], "From").click()

# to
report_to = report_to.split()
driver.find_element(By.XPATH, '//*[@id="ctl00_cphPage_cebc_Date_RangeToDate_down"]').click()
dropdown = driver.find_element(By.XPATH, '//*[@id="ctl00_cphPage_cebc_Date_RangeToDate__Calendar_YearsDropDown"]')
dd = Select(dropdown)
dd.select_by_visible_text(report_to[2])
dropdown = driver.find_element(By.XPATH, '//*[@id="ctl00_cphPage_cebc_Date_RangeToDate__Calendar_MonthsDropDown"]')
dd = Select(dropdown)
dd.select_by_visible_text(report_to[1])
get_date_element(report_to[0], "To").click()

# download
print("Please wait till download ...")
driver.find_element(By.XPATH, '//*[@id="ctl00_pageToolbar_I8"]/table').click()

# success
print("Download successful!")

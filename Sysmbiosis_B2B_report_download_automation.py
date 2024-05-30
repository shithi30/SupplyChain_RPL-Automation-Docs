# import
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import time
from datetime import date, timedelta

# preferences
options=webdriver.ChromeOptions()
options.add_argument("start-maximized")
options.add_argument("disable-infobars")
options.add_argument("--disable-extensions")

# URL
driver=webdriver.Chrome(options=options)
driver.get("http://app.ubldms.com/OnlineSales/Default")

# login
driver.find_element(By.LINK_TEXT, 'Sign in').click()

# user
driver.find_element(By.XPATH, '//*[@id="ctl00_MainContent_txtLoginID"]').send_keys('Rashedul')

# pass (spell out chain)
elem=driver.find_element(By.XPATH, '//*[@id="ctl00_MainContent_txtPasswordMask"]')
actionChains=ActionChains(driver)
actionChains.move_to_element(elem).click().send_keys("Rokon01911@#!").perform()

# submit
driver.find_element(By.XPATH, '//*[@id="MainContent_btnOK"]').click()

# Reports > All DMS Reports > Symbiosis Report Request
driver.find_element(By.LINK_TEXT, 'Reports').click()
time.sleep(1)
driver.find_element(By.LINK_TEXT, 'All DMS Reports').click()
time.sleep(1)
driver.find_element(By.LINK_TEXT, 'Symbiosis Reports Request').click()

# yesterday
elem=driver.find_element(By.XPATH, '//*[@id="ctl00_MainContent_dtpStartDate_dateInput"]')
elem.clear()
elem.send_keys((date.today()-timedelta(days=1)).strftime('%d/%m/%y')+'\n')
time.sleep(5)

# processed
elem=driver.find_element(By.XPATH, '//*[@id="ctl00_MainContent_rgvRQTrans_ctl00__0"]/td[3]')
actionChains=ActionChains(driver)
actionChains.move_to_element(elem).double_click(elem).perform()
time.sleep(1)

# download
driver.switch_to.frame(driver.find_element(By.XPATH, '//*[@id="RadWindowWrapper_ctl00_MainContent_ROITran"]/table/tbody/tr[2]/td[2]/iframe'))
elem=driver.find_element(By.XPATH, '//*[@id="btnDownload"]')
elem.click()

# success
print("Click successful, wait till full download.")

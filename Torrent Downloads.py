import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

driver = webdriver.Safari()
driver.get("http://192.168.0.85:8181/tos/index.php?user/login")
driver.maximize_window()
username = driver.find_element_by_id("username")
username.clear()
time.sleep(1)
username.send_keys("admin")
password = driver.find_element_by_id("password")
password.clear()
password.send_keys("Passw0rd!")
driver.find_element_by_id('submit').click()
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

driver = webdriver.Safari()
driver.get("https://unmineable.com/coins/ADA/address/addr1qxgwsgc83v3l2vnzey0c3kvu4lfmxewzyvglj0p0yrz6urw6wycg7a49wc6quf5a6fjtaktuepvd94pljzncpuyd057sl6dvpj")
driver.maximize_window()
search_bar = driver.find_element_by_id("address_input")
search_bar.clear()
search_bar.send_keys("addr1qxgwsgc83v3l2vnzey0c3kvu4lfmxewzyvglj0p0yrz6urw6wycg7a49wc6quf5a6fjtaktuepvd94pljzncpuyd057sl6dvpj")
value = driver.find_element_by_id("pending_balance")
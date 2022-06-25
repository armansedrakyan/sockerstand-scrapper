from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import NoSuchElementException, ElementNotVisibleException, ElementNotSelectableException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import requests
import time
import csv
import os
import logging
from fuzzywuzzy import fuzz
from datetime import datetime
import re
import math
import xlsxwriter

PATH ="C:\Program Files\Google\Chrome\chromedriver.exe";
driver=webdriver.Chrome(PATH);
driver.maximize_window()
driver.get("https://www.soccerstand.com/ru/");
time.sleep(2)
driver.find_element("css selector","#g_1_INstta7D").click()
time.sleep(5)
driver.switch_to.window(driver.window_handles[1])
driver.find_element("css selector","#detail > div.tabs.tabs__detail > div > a:nth-child(2)").click()
time.sleep(2)                
win1 = driver.find_element("css selector","#detail > div:nth-child(7) > div.oddsTab__tableWrapper > div > div.ui-table__body > div:nth-child(1) > a:nth-child(2) > span").text
win0 = driver.find_element("css selector","#detail > div:nth-child(7) > div.oddsTab__tableWrapper > div > div.ui-table__body > div:nth-child(1) > a:nth-child(3) > span").text
win2 = driver.find_element("css selector","#detail > div:nth-child(7) > div.oddsTab__tableWrapper > div > div.ui-table__body > div:nth-child(1) > a:nth-child(4) > span").text
driver.close()
# switch back to old window with switch_to.window()
driver.switch_to.window(driver.window_handles[0])
print(win1)
print(win0)
print(win2)
driver.find_element("css selector","#g_1_INstta7D").click()
time.sleep(5)
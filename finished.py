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
driver.find_element("css selector",".filters__tab:nth-child(3)").click();
time.sleep(5)
jobHeader=['day','Time','Ligue','Game','P1', 'X'];
jobBody=[];
ligue="";
day="";
for i in range(1):
    i+=1
    if i > 0:
        driver.find_element("css selector",".calendar__navigation.calendar__navigation--yesterday").click();
        time.sleep(5)
    day=driver.find_element("css selector",".calendar__datepicker").text;
    index=0;
    if index==5:
        break;
    for item in driver.find_elements("css selector",".sportName > div"):
        index=index+1;
        if item.get_attribute('class').split(" ")[0] == 'event__header':
            ligue=driver.find_element("css selector","#live-table > div.event.odds > div > div > div:nth-child("+str(index)+") > div:nth-child(2) > div:nth-child(1) > span:nth-child(1)").text+" "+driver.find_element("css selector","#live-table > div.event.odds > div > div > div:nth-child("+str(index)+") > div:nth-child(2) > div:nth-child(1) > span:nth-child(1)").text
        else:
            timex=driver.find_element("css selector","#live-table > div.event.odds > div > div > div:nth-child("+str(index)+") > div:nth-child(2)").text;
            home=driver.find_element("css selector","#live-table > div.event.odds > div > div > div:nth-child("+str(index)+") > div.event__participant--home").text;
            away=driver.find_element("css selector","#live-table > div.event.odds > div > div > div:nth-child("+str(index)+") > div.event__participant--away").text;
            
            win1 = driver.find_element("css selector","#live-table > div.event.odds > div > div > div:nth-child("+str(index)+") > div.event__score.event__score--home").text
            win0 = driver.find_element("css selector","#live-table > div.event.odds > div > div > div:nth-child("+str(index)+") > div.event__score.event__score--away").text
            jobBody.append([day,timex,ligue,home+"-"+away,win1,win0]);
         
            #break;
        if index == 5:
            break;
#with open('soccer.csv', 'a',newline='',encoding="utf8") as f:
    # using csv.writer method from CSV package
#    write = csv.writer(f)
#    write.writerow(jobHeader)
#    write.writerows(jobBody)  
    
workbook = xlsxwriter.Workbook('data/finished.xlsx')
worksheet = workbook.add_worksheet()

row = 0
col = 0

colCount=0
for item in (jobHeader):
    worksheet.write(0, colCount, item)
    colCount+=1
rowCount=1
for day,timex,ligue,home,win1,win0 in (jobBody):
    worksheet.write(rowCount, 0, day)
    worksheet.write(rowCount, 1, timex)
    worksheet.write(rowCount, 2, ligue)
    worksheet.write(rowCount, 3, home)
    worksheet.write(rowCount, 4, win1)
    worksheet.write(rowCount, 5, win0)
    rowCount+=1
# Iterate over the data and write it out row by row.

# for col in worksheet.columns:
     # max_length = 0
     # column = col[0].column_letter # Get the column name
     # for cell in col:
         # try: # Necessary to avoid error on empty cells
             # if len(str(cell.value)) > max_length:
                 # max_length = len(str(cell.value))
         # except:
             # pass
     # adjusted_width = (max_length + 2) * 1.2
     # worksheet.column_dimensions[column].width = adjusted_width
    
workbook.close()
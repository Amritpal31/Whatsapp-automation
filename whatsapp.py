from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import sys
import openpyxl as excel
import pyautogui as gui
import os


driver = webdriver.Chrome('./chromedriver')

def readContacts(fileName):
    lst = []
    file = excel.load_workbook(fileName)
    sheet = file.active
    firstCol = sheet['A']
    for cell in range(len(firstCol)):
        contact = (firstCol[cell].value)
        lst.append(contact)
    return lst


targets = readContacts("Test.xlsx")


driver.get("https://web.whatsapp.com/")
driver.maximize_window()
time.sleep(10)
wait = WebDriverWait(driver, 600)
wait5 = WebDriverWait(driver, 5)
msg = ["Hello"," ","I am sending this message using python script", " ", "Regards"]


for target in range(0,5):
    search_box=driver.find_element_by_xpath("//*[@id='side']/div[1]/div/label/div/div[2]")
    search_box.send_keys(targets[target])
    #time.sleep(2)
    try:
        user = driver.find_element_by_xpath("//span[@title='{}']".format(targets[target]))
        user.click()

        for item in msg:
            msg_box = driver.find_element_by_xpath("//*[@id='main']/footer/div[1]/div[2]/div/div[2]")
            msg_box.send_keys(item)
            gui.hotkey('shift', 'enter')
        #time.sleep(1)
        driver.find_element_by_xpath("//*[@id='main']/footer/div[1]/div[3]/button").click()
        #time.sleep(2)
    except:
        gui.hotkey('ctrl','a')
        gui.press('backspace')
        continue
time.sleep(2)
driver.close()

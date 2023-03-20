from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import calendar
import time
from datetime import datetime 
import os
import openpyxl
from openpyxl import Workbook
from xlsxwriter.workbook import Workbook
import tkinter as tk
import tkinter.messagebox as msgbox
import pyautogui
import win32com.client as win32
from pywintypes import com_error
import win32com.client
import fitz
import shutil

#Variables
FILENAME = "XXX"
NEW_NAME = "XXX.xlsx"
current_GMT = time.gmtime()
USERS = os.path.expanduser('~')
FOLDERNAME='Extract'
Downloads=(f'{USERS}\Downloads')
Documents=(f'{USERS}\Documents')
Extract = (f'{Documents}\{FOLDERNAME}')
REPORT = "REPORT.xlsx"
emailOutlook = 'almog.davidson@eiffage.com'
passwordOutlook = 'Valens2022'



#Webdriver config:
service_obj = Service("WebDrivers_path\chromedriver.exe")
driver = webdriver.Chrome(service=service_obj)

#Outlook

driver.maximize_window()
driver.get("https://outlook.office365.com/mail/AAMkADNjMDFhMDljLTcxNDItNDU1My04ZWJkLWE3MjY3YzQyMWE4NgAuAAAAAABtZmj7uFKxTJUDw1rK%2B2UaAQClX%2BsphslyRLVqrG8fet%2FsAAAQ6hlfAAA%3D")
time.sleep(20)

driver.find_element(By.XPATH, '/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[2]/div[2]/div/input[1]').send_keys(emailOutlook)
time.sleep(5)
driver.find_element(By.XPATH, '/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[4]/div/div/div/div/input').click()
time.sleep(5)

driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[1]/div[2]/div/div/form/div[2]/div[2]/input').send_keys(passwordOutlook)
time.sleep(5)
driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[1]/div[2]/div/div/form/div[2]/div[4]/span').click()
time.sleep(5)
driver.find_element(By.XPATH, '/html/body/div/form/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[3]/div[2]/div/div/div[1]/input').click()
time.sleep(20)

# BIM360 Ribbon

driver.find_element(By.XPATH,)

driver.find_element(By.XPATH, '/html/body/div[2]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input[1]').send_keys('BIM360')
time.sleep(5) 



driver.find_element(By.XPATH, '/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div[3]/div/div[3]/div[1]/div[2]/div/div/div/div/div/div[1]/div/div/div[1]/div[1]/div').click()
time.sleep(20)


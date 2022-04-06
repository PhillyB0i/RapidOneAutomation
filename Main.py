################################################################
#############      Created by Philip Boubenchikov     ##########
#############              For RapidOne               ##########
#############               Build 1.0.0               ##########
################################################################
from multiprocessing import parent_process
from multiprocessing.connection import Listener, wait
from multiprocessing.sharedctypes import Value
import os
from platform import release
from sched import scheduler
from sre_constants import JUMP
import tkinter
from xml.etree.ElementTree import TreeBuilder
from numpy import number
import requests # pip install requests
import zipfile
import pynput
from pynput.keyboard import Listener
import sys
import tkinter as tk
from tkinter import *
import tkinter.messagebox
import subprocess
import psutil
import logging
from datetime import datetime
import linecache
import time
from requests.sessions import dispatch_hook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoAlertPresentException
from time import sleep
from pynput import keyboard
from pynput.keyboard import Key, Controller
from selenium.webdriver.chrome.options import Options

current = []
#Chromedriver URL, choices, change to cwd path
Durl = 'https://chromedriver.storage.googleapis.com/99.0.4844.51/chromedriver_win32.zip'
chromedriver = requests.get(Durl)
FileName = chromedriver.url[Durl.rfind('/')+1:]
path = os.getcwd()
DriverExist = 'chromedriver.exe'
PathDriver = os.path.isfile(DriverExist)

def clearConsole():
    command = 'clear'
    if os.name in ('nt', 'dos'):  # If Machine is running on Windows, use cls
        command = 'cls'
    os.system(command)


print("Made by Philip Boubenchikov"'\n')
print('Build 1.0.0')

chromeoptions = Options()
chromeoptions.add_argument("--log-level=3")

while True:

#Opens ONE
    if PathDriver is False:
            print('[WARNING]: Chromedriver file is missing..')
            with open(FileName, 'wb') as f:
                for chunk in chromedriver.iter_content(chunk_size=2561100122314):
                    if chunk:
                        f.write(chunk)
                        print('Downloading...')
                        print('Extracting...')
                        sleep(5)
                    with zipfile.ZipFile('chromedriver_win32.zip', 'r') as zip_ref:
                        zip_ref.extract('chromedriver.exe')
                        sleep(1)
    break
if PathDriver is False:
    print('Done. Restart software Please.')
    print('Press CNTRL + C to exit.')
    for remaining in range(10, -1, -1):
        sys.stdout.write("\r")
        sys.stdout.write("In{:2d} Seconds Program will exit".format(remaining))
        sys.stdout.flush()
        time.sleep(1)
        if remaining == 0:
            exit()
        
elif PathDriver is True:
        while True:
            print("Opening Rapid..")
            sleep(1)
            exists = 'Config.ini'
            os_isfile = os.path.isfile(exists)
            EmailPaste = linecache.getline('Config.ini', 1)
            PassPaste = linecache.getline('Config.ini', 2)
            WebsitePaste = linecache.getline('Config.ini', 3)
            RapidURL = (R"https://" + WebsitePaste + ".rapid-image.net")
            
            if os_isfile is True:
                driver = webdriver.Chrome(options=chromeoptions)
                driver.get(RapidURL)
                sleep(2)
                EmailField = driver.find_element(by=By.XPATH, value="//input[@type='text']")
                Do = ActionChains(driver)
                #Do.click(EmailField)
                EmailField.send_keys(EmailPaste)
                PasswordField = driver.find_element(by=By.XPATH, value="//input[@type='password']")
                #Do.click(PasswordField)
                PasswordField.send_keys(PassPaste)
                SignIn = driver.find_element(by=By.XPATH, value="//button[@class='login_btn btn ng-binding']")
                Do.click(SignIn)
                break
                

            if os_isfile is False:
                print("[WARNING]: Login Config file missing please type login and site info:")
                print()
            Email = input("UserName:")
            Password = input("Password:")
            Website = input("Rapid URL - XXXX.rapid-image.net - only XXXX part:")
            with open('Config.ini', 'w') as f:
                f.write('{}{}{}'.format(Email,'\n'+Password,'\n'+Website))

sleep(5)
clearConsole()
print('['+datetime.now().strftime("%H:%M:%S")+']:',"Press CONTROL + F1 to show list of shortcuts")

#HotKeys
#General
ShortcutBOX = [keyboard.Key.ctrl_l, keyboard.Key.f1]
FinDoc = [keyboard.Key.ctrl_l, keyboard.Key.f2]
LastPatient = [keyboard.Key.ctrl_l, keyboard.Key.f3]
scheduletoday = [keyboard.Key.ctrl_l, keyboard.Key.f6]
#Reports
treatmentreport = [keyboard.Key.ctrl_l, keyboard.Key.f7]
patientreport = [keyboard.Key.ctrl_l, keyboard.Key.f8]
attendanceman = [keyboard.Key.ctrl_l, keyboard.Key.f9]
dailyincome = [keyboard.Key.ctrl_l, keyboard.Key.f10]
#Exit
exitscript = [keyboard.Key.alt_l, keyboard.Key.esc]
#CRM
leadmanagment = [keyboard.Key.alt_l, keyboard.Key.f1]
movenext = [keyboard.Key.ctrl_l, keyboard.Key.right]
moveback = [keyboard.Key.ctrl_l, keyboard.Key.left]
jumpfirst = [keyboard.Key.ctrl_l, keyboard.Key.up]
jumplast = [keyboard.Key.ctrl_l, keyboard.Key.down]
showlatest0 = [keyboard.Key.alt_l, keyboard.Key.f2]
showlatest7 = [keyboard.Key.alt_l, keyboard.Key.f3]
showlatest14 = [keyboard.Key.alt_l, keyboard.Key.f5]

#Popup ShortCut 
def executeMenu():
    print('['+datetime.now().strftime("%H:%M:%S")+']:',"Showing shortcuts list")
    tkinter.messagebox.showinfo("Rapid Shortcuts",  "General:"'\n'"Hotkey List: LEFT CONTROL + F1"'\n'"Create Financial Document: LEFT CONTROL + F2"'\n'"Jump to latest patient: LEFT CONTROL + F3"'\n'"Switch schedule to today: LEFT CONTROL + F6"'\n''\n'"Reports:"'\n'"Jump to Treatment Report: LEFT CONTROL + F7"'\n'"Jump to Patient Report: LEFT CONTROL + F8"'\n'"Jump to Attendance Managment: LEFT CONTROL + F9"'\n'"Jump to Custom daily income report: LEFT CONTROL + F10"'\n''\n'"CRM:"'\n'"Move to lead managment: LEFT ALT + F1"'\n'"Move to next page: LEFT CONTROL + Right arrow"'\n'"Move to previous page: CONTROL ALT + Left arrow"'\n'"Jump to first page: LEFT CONTROL + Up arrow"'\n'"Jump to last page: LEFT CONTROL + Down arrow"'\n'"Show latest leads for today: LEFT ALT + F2"'\n'"Show latest leads for last 7 days: LEFT ALT + F3"'\n'"Show latest leads for last 14 days: LEFT ALT + F5")
    
#Move to FinDoc
def executeFinDoc():
    print ('['+datetime.now().strftime("%H:%M:%S")+']:',"Selected Create Financial Document")
    driver.get(RapidURL+"/financial/new-sale")

#Move to latest patient visited
def executeLastPat():
    print ('['+datetime.now().strftime("%H:%M:%S")+']:',"Selected Latest Patient")
    driver.find_element(by=By.XPATH, value=u"//img[@title='לקוחות קודמים']").click()
    sleep(1)
    driver.find_element(by=By.XPATH, value=u"//li/div/div/a").click()

#Jump to current day in schedule
def executescheduletoday():
    if "/schedule" in driver.current_url:
        print ('['+datetime.now().strftime("%H:%M:%S")+']:',"Jumping to today in schedule")
        driver.find_element(by=By.XPATH, value=u"//a[contains(text(),'היום')]").click()
    elif "/schedule" not in driver.current_url:
        print('['+datetime.now().strftime("%H:%M:%S")+']',"[WARNING]: Not in schedule, please go to schedule")

####################################
#         Reports Section          #
####################################

#Jump to treatment report
def executetreatmentreport():
    print ('['+datetime.now().strftime("%H:%M:%S")+']:', "Opening treatment report")
    driver.get(RapidURL+"/reports/treatments-report")

#Jump to patient report
def executepatientreport():
    print ('['+datetime.now().strftime("%H:%M:%S")+']:', "Opening patient report")
    driver.get(RapidURL+"/reports/patients")

#Jump to attendance management
def executeattendance():
    print ('['+datetime.now().strftime("%H:%M:%S")+']:', "Opening attendance management")
    driver.get(RapidURL+"/reports/attendance-management")

#Create daily income report
def executetodayincome():
    print ('['+datetime.now().strftime("%H:%M:%S")+']:', "Creating daily income report")
    driver.get(RapidURL+"/reports/income-report")
    sleep(2)
    driver.find_element(by=By.XPATH, value=u"//input[@type='search']").click()
    sleep(1)
    driver.find_element(by=By.XPATH, value=u"//*/text()[normalize-space(.)='הכל']/parent::*").click()
    sleep(1)
    driver.find_element(by=By.XPATH, value=u"//form[@id='search_form']/div[2]/div/div/span/span/span/span").click()
    sleep(1)
    driver.find_element(by=By.CSS_SELECTOR, value=u"a.k-link.k-nav-today").click()
    sleep(1)
    driver.find_element(by=By.ID, value=u"search-btn").click()

####################################
#      End Reports Section         #
####################################

####################################
#          CRM Section             #
####################################

#Jump to Lead managment
def executeleadmanamgent():
    print ('['+datetime.now().strftime("%H:%M:%S")+']:', "Jumping to lead managment")
    driver.get(RapidURL+"/crm/lead-management/")

#Move to the next lead page
def executemovenext():
    if "/crm/lead-management/" in driver.current_url:
        print ('['+datetime.now().strftime("%H:%M:%S")+']:',"Moving to next leads page")
        driver.find_element(by=By.XPATH, value=u"//*/text()[normalize-space(.)='›']/parent::*").click()
    elif "/crm/lead-management/" not in driver.current_url:
        print('['+datetime.now().strftime("%H:%M:%S")+']',"[WARNING]: Not in CRM please go to CRM to use this shortcut")

#Move back to previous leads page
def executemoveback():
    if "/crm/lead-management/" in driver.current_url:
        print ('['+datetime.now().strftime("%H:%M:%S")+']:',"Moving to previous leads page")
        driver.find_element(by=By.XPATH, value=u"//*/text()[normalize-space(.)='‹']/parent::*").click()
    elif "/crm/lead-management/" not in driver.current_url:
        print('['+datetime.now().strftime("%H:%M:%S")+']',"[WARNING]: Not in CRM please go to CRM to use this shortcut")

#Jump to first page
def executejumpfirst():
    if "/crm/lead-management/" in driver.current_url:
        print ('['+datetime.now().strftime("%H:%M:%S")+']:',"Jumping to first page")
        driver.find_element(by=By.XPATH, value=u"//*/text()[normalize-space(.)='ראשון']/parent::*").click()
    elif "/crm/lead-management/" not in driver.current_url:
        print('['+datetime.now().strftime("%H:%M:%S")+']',"[WARNING]: Not in CRM please go to CRM to use this shortcut")

#Jump to last page
def executejumplast():
    if "/crm/lead-management/" in driver.current_url:
        print ('['+datetime.now().strftime("%H:%M:%S")+']:',"Jumping to last page")
        driver.find_element(by=By.XPATH, value=u"//*/text()[normalize-space(.)='אחרון']/parent::*").click()
    elif "/crm/lead-management/" not in driver.current_url:
        print('['+datetime.now().strftime("%H:%M:%S")+']',"[WARNING]: Not in CRM please go to CRM to use this shortcut")

#Show leads only for today
def executeshowlatest0():
    if "/crm/lead-management/" in driver.current_url:
        print ('['+datetime.now().strftime("%H:%M:%S")+']:',"Showing leads for today")
        driver.find_element(by=By.XPATH, value=u"//button[@type='button']").click()
        driver.find_element(by=By.XPATH, value=u"(.//*[normalize-space(text()) and normalize-space(.)='סוג תאריך'])[1]/following::select[1]").click()
        Select(driver.find_element(by=By.XPATH, value=u"(.//*[normalize-space(text()) and normalize-space(.)='סוג תאריך'])[1]/following::select[1]")).select_by_visible_text(u"גיל הליד")
        driver.find_element(by=By.ID, value=u"age").click()
        Select(driver.find_element(by=By.ID, value=u"age")).select_by_visible_text("0")
        driver.find_element(by=By.XPATH, value=u"//*/text()[normalize-space(.)='הפעל']/parent::*").click()
    elif "/crm/lead-management/" not in driver.current_url:
        print('['+datetime.now().strftime("%H:%M:%S")+']',"[WARNING]: Not in CRM please go to CRM to use this shortcut")

#Show leads past 7 days
def executeshowlatest7():
    if "/crm/lead-management/" in driver.current_url:
        print ('['+datetime.now().strftime("%H:%M:%S")+']:',"Showing leads of the past 7 days")
        driver.find_element(by=By.XPATH, value=u"//button[@type='button']").click()
        driver.find_element(by=By.XPATH, value=u"(.//*[normalize-space(text()) and normalize-space(.)='סוג תאריך'])[1]/following::select[1]").click()
        Select(driver.find_element(by=By.XPATH, value=u"(.//*[normalize-space(text()) and normalize-space(.)='סוג תאריך'])[1]/following::select[1]")).select_by_visible_text(u"גיל הליד")
        driver.find_element(by=By.ID, value=u"age").click()
        Select(driver.find_element(by=By.ID, value=u"age")).select_by_visible_text("7")
        driver.find_element(by=By.XPATH, value=u"//*/text()[normalize-space(.)='הפעל']/parent::*").click()
    elif "/crm/lead-management/" not in driver.current_url:
        print('['+datetime.now().strftime("%H:%M:%S")+']',"[WARNING]: Not in CRM please go to CRM to use this shortcut")

#Show leads past 14 days
def executeshowlatest14():
    if "/crm/lead-management/" in driver.current_url:
        print ('['+datetime.now().strftime("%H:%M:%S")+']:',"Showing leads of the past 14 days")
        driver.find_element(by=By.XPATH, value=u"//button[@type='button']").click()
        driver.find_element(by=By.XPATH, value=u"(.//*[normalize-space(text()) and normalize-space(.)='סוג תאריך'])[1]/following::select[1]").click()
        Select(driver.find_element(by=By.XPATH, value=u"(.//*[normalize-space(text()) and normalize-space(.)='סוג תאריך'])[1]/following::select[1]")).select_by_visible_text(u"גיל הליד")
        driver.find_element(by=By.ID, value=u"age").click()
        Select(driver.find_element(by=By.ID, value=u"age")).select_by_visible_text("14")
        driver.find_element(by=By.XPATH, value=u"//*/text()[normalize-space(.)='הפעל']/parent::*").click()
    elif "/crm/lead-management/" not in driver.current_url:
        print('['+datetime.now().strftime("%H:%M:%S")+']',"[WARNING]: Not in CRM please go to CRM to use this shortcut")

####################################
#        End CRM Section           #
####################################

#Kill Chromedriver and chrome and exit script
def exitsoft():
    subprocess.call("TASKKILL /f  /IM  CHROME.EXE")
    subprocess.call("TASKKILL /f  /IM  CHROMEDRIVER.EXE")
    sleep(2)
    listener.stop()
    exit()

#execute Shortcut INFO
def HotkeyPress(key):
    global current
    current.append(key)
    if(len(current) == 2):
        if(current[0] == ShortcutBOX[0] and current[1] == ShortcutBOX[1]):
            executeMenu()
        elif (current[0] == FinDoc[0] and current[1] == FinDoc[1]):
            executeFinDoc()
        elif (current[0] == LastPatient[0] and current[1] == LastPatient[1]):   
            executeLastPat()
        elif (current[0] == exitscript[0] and current[1] == exitscript[1]):   
            exitsoft()
        elif (current[0] == scheduletoday[0] and current[1] == scheduletoday[1]):
            executescheduletoday()
        elif (current[0] == treatmentreport[0] and current[1] == treatmentreport[1]):
            executetreatmentreport()
        elif (current[0] == patientreport[0] and current[1] == patientreport[1]):
            executepatientreport()
        elif (current[0] == attendanceman[0] and current[1] == attendanceman[1]):
            executeattendance()
        elif (current[0] == dailyincome[0] and current[1] == dailyincome[1]):
            executetodayincome()
        elif (current[0] == leadmanagment[0] and current[1] == leadmanagment[1]):
            executeleadmanamgent()
        elif (current[0] == movenext[0] and current[1] == movenext[1]):
            executemovenext()
        elif (current[0] == moveback[0] and current[1] == moveback[1]):
            executemoveback()
        elif (current[0] == jumpfirst[0] and current[1] == jumpfirst[1]):
            executejumpfirst()
        elif (current[0] == jumplast[0] and current[1] == jumplast[1]):
            executejumplast()
        elif (current[0] == showlatest14[0] and current[1] == showlatest14[1]):
            executeshowlatest14()
        elif (current[0] == showlatest7[0] and current[1] == showlatest7[1]):
            executeshowlatest7()
        elif (current[0] == showlatest0[0] and current[1] == showlatest0[1]):
            executeshowlatest0()
        current = []

#Listen for keyboard presses
with keyboard.Listener(on_press=HotkeyPress) as listener:
    listener = Listener()
    listener.start()
    listener.join()


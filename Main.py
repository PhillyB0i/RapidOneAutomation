################################################################
#############      Created by Philip Boubenchikov     ##########
#############              For RapidOne               ##########
#############               Build 0.0.1               ##########
################################################################
from multiprocessing.connection import wait
from multiprocessing.sharedctypes import Value
import os
from sched import scheduler
from xml.etree.ElementTree import TreeBuilder
from numpy import number
import requests # pip install requests
import zipfile
import pynput
import sys
import tkinter as tk
from tkinter import *
import tkinter.messagebox
import subprocess
import psutil
import logging
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
from time import sleep
from pynput import keyboard

#Chromedriver URL, choices, change to cwd path
Durl = 'https://chromedriver.storage.googleapis.com/99.0.4844.51/chromedriver_win32.zip'
chromedriver = requests.get(Durl)
yesChoice = ['y', 'Y', 'ye', 'Ye', 'Yes', 'yes', 'YES', 'YE']
noChoice = ['n', 'N', 'no', 'No', 'Nah', 'nah', 'NAH', 'NO']
FileName = chromedriver.url[Durl.rfind('/')+1:]
path = os.getcwd()
DriverExist = 'chromedriver.exe'
PathDriver = os.path.isfile(DriverExist)


def clearConsole():
    command = 'clear'
    if os.name in ('nt', 'dos'):  # If Machine is running on Windows, use cls
        command = 'cls'
    os.system(command)


print('Made by Philip Boubenchikov')
print('Build 0.0.1')
while True:

#Opens ONE
    if PathDriver is False:
            print('Chromedriver file is missing..')
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
    for remaining in range(10, 0, -1):
        sys.stdout.write("\r")
        sys.stdout.write("{:2d} Program will exit".format(remaining))
        sys.stdout.flush()
        time.sleep(1)

elif PathDriver is True:
        while True:
            print("Opening Rapid..")
            sleep(1)
            exists = 'Details.txt'
            os_isfile = os.path.isfile(exists)
            EmailPaste = linecache.getline('Details.txt', 1)
            PassPaste = linecache.getline('Details.txt', 2)
            WebsitePaste = linecache.getline('Details.txt', 3)
            RapidURL = (R"https://" + WebsitePaste + ".rapid-image.net")
            
            if os_isfile is True:
                driver = webdriver.Chrome('chromedriver.exe')
                driver.get(RapidURL)
                sleep(2)
                EmailField = driver.find_element(by=By.XPATH, value="//input[@placeholder='Username']")
                Do = ActionChains(driver)
                #Do.click(EmailField)
                EmailField.send_keys(EmailPaste)
                PasswordField = driver.find_element(by=By.XPATH, value="//input[@placeholder='Password']")
                #Do.click(PasswordField)
                PasswordField.send_keys(PassPaste)
                SignIn = driver.find_element(by=By.XPATH, value="//button[@class='login_btn btn ng-binding']")
                Do.click(SignIn)
                break
                

            if os_isfile is False:
                print("Login Details file missing please type login info:")
                print()
            Email = input("UserName:")
            Password = input("Password:")
            Website = input("Rapid URL - XXXX.rapid-image.net - only XXXX part:")
            with open('Details.txt', 'w') as f:
                f.write('{}{}{}'.format(Email,'\n'+Password,'\n'+Website))

sleep(5)
clearConsole()
print('Press F1 to show list of shortcuts')

#HotKeys
ShortcutBOX = [
    {keyboard.Key.f1}
]
FinDoc = [
    {keyboard.Key.f2}
]

# The currently active modifiers
current = set()

#Popup ShortCut 
def executeMenu():
    print('Showing ShortCut List')
    tkinter.messagebox.showinfo("Rapid Shortcuts",  "F1 - Shortcut list"'\n'"F2 - Create Financial Document")
    clearConsole()


#execute on key press Shortcut INFO
def ShortcutBOXMenu(key):
    if any([key in COMBO for COMBO in ShortcutBOX]):
        current.add(key)
        if any(all(k in current for k in COMBO) for COMBO in ShortcutBOX):
            executeMenu()    

#Listen for keyboard presses
with keyboard.Listener(on_press=ShortcutBOXMenu) as listener:
    listener.join()


    #$$$ Doc execution
    def executeFinDoc():
        print ("Selected Create Financial Document")
        driver.get(RapidURL+"/financial/new-sale")

    #execute on key press $$ Doc
    def Finny(key):
        if any([key in COMBO for COMBO in FinDoc]):
            current.add(key)
            if any(all(k in current for k in COMBO) for COMBO in FinDoc):
                executeFinDoc()
    with keyboard.Listener(on_press=Finny) as listener:
        listener.join()
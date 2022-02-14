
pwOR = "a0mvdt2d"
usrNameOR = "johnson.bolaji@commscope.com"

#import required libraries

from openpyxl import Workbook, load_workbook, drawing
from openpyxl.drawing import line, image
from openpyxl.drawing.image import Image
from PIL import Image
from io import BytesIO
from datetime import date
#import tkinter as tk
import easygui
import subprocess
import os
from easygui import *
import selenium
import shutil
import glob
import os.path
import getpass
import cv2
import numpy as np
import choco
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import chromedriver_autoinstaller
import time


#----------------------------------------------------------------

class Pole:
    def __init__(BT_ID,xGRF,yGRF,AFN1,AFN2):

        BT_ID.xGRF = xGRF
        BT_ID.yGRF = yGRF
        BT_ID.AFN1 = AFN1
        BT_ID.AFN2 = AFN2


FILEBROWSER_PATH = os.path.join(os.getenv('WINDIR'), 'explorer.exe')

def explore(path):
    # explorer would choke on forward slashes
    path = os.path.normpath(path)

    if os.path.isdir(path):
        subprocess.run([FILEBROWSER_PATH, path])
    elif os.path.isfile(path):
        subprocess.run([FILEBROWSER_PATH, '/select,', os.path.normpath(path)])

user = getpass.getuser() #get windows username

#config = configparser.ConfigParser()

chromedriver_autoinstaller.install()

print("Select folder containing A55 documents. Note: please make sure folder only contains A55 files that you wish to edit.")

A55Path = easygui.diropenbox()

print("Select folder to contain your screenshots")

screenshotPath = easygui.diropenbox()

for filename in os.listdir(A55Path):
    if filename.endswith(".xlsx"):

        substring = "A55"

        A55Book = load_workbook(A55Path + '\\' + filename)

        #check containing A55 text
        
        for s in range(len(A55Book.sheetnames)):
            if substring in A55Book.sheetnames[s]:
                break
        A55Book.active = s

        cover = A55Book.active

        fullGRF = cover['O4'].internal_value

        GRF = fullGRF.split(",")
        GRF[0] = GRF[0].strip()
        GRF[1] = GRF[1].strip()

        workPole = Pole(GRF[0],GRF[1],'xox','xoxx')
        #-------------------------------------------- get SS
        #PIAscreenshot(workPole.xGRF,workPole.yGRF,cdPATH,usrNameOR,pwOR) #call screenshot function, NB put details in config file

        driver = webdriver.Chrome()

        # Open the website
        driver.get('https://www.openreach.co.uk/cpportal/login')

        #element = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((CSS_SELECTOR, ".panel-title a")))
        #element.click()

        id_box = driver.find_element_by_xpath("//*[@id='smLoginFormId']/div[1]/div[1]/input")

        id_box.click

        id_box.send_keys(usrNameOR)

        pw_box = driver.find_element_by_xpath("//*[@id='smLoginFormId']/div[1]/div[2]/input")

        pw_box.click
    
        pw_box.send_keys(pwOR)

        login_button = driver.find_element_by_xpath("//*[@id='smLoginFormId']/div[2]/button")

        driver.implicitly_wait(5) # seconds

        pw_box.send_keys(Keys.RETURN)

        time.sleep(5)

        driver.get("https://www.openreach.co.uk/ormaps/pia/v2")

        time.sleep(5)

        threelines = driver.find_element_by_xpath("//*[@id='searchpanel']/div[1]/i")
        hover = ActionChains(driver).move_to_element(threelines).click(threelines)
        hover.perform()

        time.sleep(5)

        GRFdrop = driver.find_element_by_xpath("//*[@id='mat-expansion-panel-header-3']/span[1]/mat-panel-title")
        hover = ActionChains(driver).move_to_element(GRFdrop).click(GRFdrop)
        hover.perform()

        time.sleep(5)

        GRFxbox = driver.find_element_by_xpath("//*[@id='mat-input-2']")
        GRFybox = driver.find_element_by_xpath("//*[@id='mat-input-3']")

        GRFxbox.click
        GRFxbox.send_keys(workPole.xGRF)

        time.sleep(2)

        GRFybox.click
        GRFybox.send_keys(workPole.yGRF)

        time.sleep(2)

        GRFybox.send_keys(Keys.RETURN)

        time.sleep(5)

        zoomInButton = driver.find_element_by_xpath("//*[@id='map']/div[3]/div[2]/div[4]/button[1]")

        hover = ActionChains(driver).move_to_element(zoomInButton).click(zoomInButton)
        hover.perform()

        time.sleep(5)

        hover.perform()

        time.sleep(3)

        threelinesPost = driver.find_element_by_xpath("//*[@id='searchpanel']/div[1]/i")
        hover = ActionChains(driver).move_to_element(threelinesPost).click(threelinesPost)
        hover.perform()

        time.sleep(3)

        downloadButton = driver.find_element_by_xpath("//*[@id='cdk-accordion-child-16']/div/div/div[1]/i")
        hover = ActionChains(driver).move_to_element(downloadButton).click(downloadButton)
        hover.perform()

        time.sleep(2)

        CME = driver.find_element_by_xpath("//*[@id='mat-dialog-0']/app-download-map-dialog/div[2]/div/button[2]/span")
        hover = ActionChains(driver).move_to_element(CME).click(CME)
        hover.perform()
        time.sleep(3)
                
     
        driver.close()

        #--------------------------------------------
        folder_path = (r'C:\Users\%s\Downloads\*' %user) #navigate to downloads

        files = glob.glob(folder_path)
        
        max_file = max(files, key=os.path.getctime)

        
        image = cv2.imread(max_file) #Read screenshot
        
        height, width, channels = image.shape #Get image dimensions
        
        window_name = 'Image' #name of image window
        
        center_coordinates = (int(width/2),int(height/2)) #coords of picture center

        radius = 30 #circle radius in pixels

        color = (0, 0, 255) #(Blue,Green,Red)

        thickness = 5 #circle thickness in pixels

        cv2.circle(image, center_coordinates, radius, color, thickness)

        cv2.imwrite(screenshotPath + "\\" + filename.removesuffix('.xlsx') + '.png', image)
        

        img = drawing.image.Image(screenshotPath + "\\" + filename.removesuffix('.xlsx') + '.png')

        cover.add_image(img,'B6')

        A55Book.save(A55Path + "\\" + filename)

        
         
        continue
    else:
        continue




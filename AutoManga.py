import os
import time
import shutil
import smtplib
import requests
import pyautogui
import subprocess
from bs4 import BeautifulSoup
from importlib import import_module
from selenium import webdriver
from email.message import EmailMessage
from selenium.webdriver.common.keys import Keys

#Variables
mangaName = 'null'
chNo = '0'
to= open("ReceiverList.ini","r").read()

#Loop for different mangas
fav = open("MangaList.ini","r").read()
numOfMangas= len(fav.split())
for i in range(0, numOfMangas):
    mangaName=(fav.split(",")[i])

    #Gettting Latest Chapter
    URL = 'https://w13.mangafreak.net/Manga/'+ mangaName
    print(URL)
    r = requests.get(URL)
    soup = BeautifulSoup(r.content, 'lxml') 
    recent_chaps = soup.find('div', class_='series_sub_chapter_list')
    latest_chap =  recent_chaps.a.text.split()[1]
    chNo = latest_chap

    #Downloading File
    driver = webdriver.Chrome()
    name = 'https://images.mangafreak.net/downloads/' + mangaName+ '_' +chNo                                               
    driver.get(name)
    time.sleep(25)

    #Moving to Manga FOlder
    LOCATION1 = 'C:/Users/Shreyas/Downloads/'+ mangaName+ '_'+chNo+'.zip'
    shutil.move(LOCATION1, "C:/Users/Shreyas/Documents/0-MangaAutomate-0")

    #Conversion to MOBI
    subprocess.Popen("C:\Program Files\Kindle Comic Converter\KCC.exe")
    time.sleep(4)
    pyautogui.click(x=1097, y=594)
    time.sleep(2)
    pyautogui.click(x=584, y=493)
    time.sleep(1)
    pyautogui.click(x=782, y=457)
    pyautogui.click(x=782, y=457)
    time.sleep(1)
    pyautogui.click(x=782, y=457)
    pyautogui.click(x=782, y=457)
    time.sleep(1)
    pyautogui.click(x=964, y=622)
    time.sleep(15)

    #Sending Mail
    msg=EmailMessage()
    msg['Subject']=mangaName + ' Chapter ' + chNo
    msg['From']='Auto Manga'
    msg['To']= to

    with open("C:/Users/Shreyas/Documents/0-MangaAutomate-0/"+ mangaName+ "_" +chNo+ ".mobi","rb") as manga:
        file_data=manga.read()
        file_name=mangaName + ' Chapter ' + chNo + '.mobi'
        msg.add_attachment(file_data,maintype="application",subtype="mobi",filename=file_name)

    server =smtplib.SMTP_SSL("smtp.gmail.com", 465)
    server.login("weeklyautomanga@gmail.com", "testaccpass123")
    server.send_message(msg)
    print(mangaName+" Manga Sent")






















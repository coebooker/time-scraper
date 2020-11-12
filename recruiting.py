from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
#from selenium.webdriver.support.ui import WebDriverWait
#from selenium.webdriver.support import expected_conditions as EC
import time
import requests
from bs4 import BeautifulSoup as bs
import datetime
from ClassStructure import *
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import pandas as pd
import xlsxwriter
def getMilesplitURL(name,state,school):
    #First part here goes to milesplit and searches
    driver = webdriver.Chrome(executable_path=r'C:\Users\Coe\Desktop\Recruiting Tool\chromedriver.exe')
    driver.get("https://"+state+".milesplit.com\search")
    search = driver.find_element_by_id("postHeaderSearch")
    search.send_keys(name)
    search.send_keys(Keys.RETURN)
    time.sleep(3)
    #selects the top profile, needs error checking to make sure it's a correct school
    #uses find elements so the len>0 statement can work if there is no link there
    runner = driver.find_elements_by_xpath('//*[@id="resultsList"]/li/div[1]')
    if len(runner) > 0:
        runner = runner[0]
        print(runner.text)
        if school in runner.text:
            runner.click()
            time.sleep(2)
            return driver.current_url
        else:
            return "x"
    else:
        return "x"
                        
#Keyword is the last word before the time/distance. String is the full string e.g. 110 Meter Hurdles 13.45, it will return 13.45
def parseEvent(keyword,string):
    string = string[string.find(keyword):]
    time = string[string.find(" "):].strip()
    return result
def getTable(url):
    header  = {'User-Agent': 'Mozilla/5.0'}
    soup = bs(requests.get(url).content, 'html.parser')
    tableLst = []
    #Converts the PR Table to a Python List
    for event, time in zip(soup.select('td.event'),
                           soup.select('td.time')):
        tableLst.append((event.text, time.text))
    return tableLst
def main(name):
    url = getMilesplitData(name)
    table = getTable(url)
def getFastestTime(tableLst):
    timeDict = dict()
    for result in tableLst:
        fieldBool = False
        event = result[0]
        time = result[1]
        if "Short Course" in event:
            next()
        #need to fix this!
        if "Shotput" in event or "Discus" in event or "Javelin" in event or "Vault" in event or "Jump" in event:
            fieldBool = True
        if event not in timeDict:
            timeDict[event] = time
        else:
            currentPR = timeDict[event]
            #If it's a running event convert the time to datetime objects so that you can compare them
            if not fieldBool:
                currentPRTime = convertToDatetime(currentPR)
                newTime = convertToDatetime(time)
                if newTime < currentPRTime:
                    timeDict[event] = time
            #If it's a field event, convert the event to inches and compare
            else:
                currentPR = getInches(timeDict[event])
                newResult = getInches(time)
                if newResult > currentPR:
                    timeDict[event] = newResult
    return timeDict
#Takes time in Feet-Inches or Meters and converts to Inches
def getInches(resultStr):
    #If it's a result measured in meters
    if "m" in resultStr:
        print("Thinks it's meters")
        #Remove the m from the string
        resultStrCopy = resultStr[:-1]
        resultFloat = float(resultStrCopy)
        resultInches = resultFloat*39.37
        resultInches = round(resultInches,2)
    #If it's a result measured in feet/inches
    else:
        print("We in here")
        dashLoc = resultStr.find("-")
        feet = int(resultStr[:dashLoc])
        feet *= 12
        inches = int(resultStr[dashLoc+1:])
        resultInches = feet + inches
    return resultInches

        
#Converts string of time to a datetime class, used for evaluation of times because strings don't compare ex 4:26.85 > 4:22.95
def convertToDatetime(timeStr):
    #Used if the time has minutes
    if ":" in timeStr:
        colonLoc = timeStr.find(":")
        minutes = int(timeStr[:colonLoc])
        periodLoc = timeStr.find(".")
        seconds = int(timeStr[colonLoc+1:periodLoc])
        milliseconds = int(timeStr[periodLoc+1:])
        return datetime.time(0,minutes,seconds,milliseconds)
    #If it's an event without minutes
    else:
        periodLoc = timeStr.find(".")
        seconds = int(timeStr[:periodLoc])
        milliseconds = int(timeStr[periodLoc+1:])
        return datetime.time(0,0,seconds,milliseconds)

def googleDrive():
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()
    drive = GoogleDrive(gauth)
    file = drive.CreateFile({'id': 'INSERT FILE ID'})
    download_mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    file.GetContentFile(file['title']+'.xlsx', mimetype=download_mimetype)
def getLstFromSheet():
    df = pd.read_excel("Men's Recruiting Mass List.xlsx", "Master List")
    lst = df.values.tolist()
    return lst
def main():
    #googleDrive()
    lst = getLstFromSheet()
    for runner in lst:
        #skip runner because it's already done
        name = runner[3]
        school = runner[4]
        state = runner[5]
        URL  = getMilesplitURL(name,state,school)
        if URL == "x":
            runner[6] = "x"
            runner[8] = "Tier 4"
        else:
            PRTable = getTable(URL)
            runner[6] = URL
            runner[9] = getFastestTime(PRTable)
    df = pd.DataFrame(lst,columns=['Status','Gender','Grad Year','Name','HS','HS State','Milesplit Link','Athletic.net Link','Tier','Personal Bests','Event Group','Academics','Cell','Cell2','Email','Address','Address 2','City','State','Zip'])
    writer = pd.ExcelWriter('Compiled Results.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1')
    writer.save()       
    

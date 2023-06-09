import datetime
import os
import time
import openpyxl
import pyautogui
import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from datetime import timedelta
from openpyxl.utils import get_column_letter
import requests
import threading
import csv



def startDriver():
    """
    Starts Chrome webdriver
    """""
    newdriver = webdriver.Chrome()
    return newdriver


def startDashboard():
    """
    Loads 'https://verkehr.aachen.de' in the webdriver and maximizes it
    """""
    url = 'https://verkehr.aachen.de'
    driver.get(url)
    driver.maximize_window()
    time.sleep(2)
    return True


def startWeather():
    """
    Returns the html source from 'https://www.wetter.com/wetter_aktuell/rueckblick/deutschland/aachen/DE0000003.html?timeframe=30d'
    """""
    url = 'https://www.wetter.com/wetter_aktuell/rueckblick/deutschland/aachen/DE0000003.html?timeframe=30d'
    response = requests.get(url)
    html_source = response.text
    return html_source


def startPictures():
    """
    Returns the html source from 'https://www.wetteronline.de/wetter/aachen'
    """""
    url = 'https://www.wetteronline.de/wetter/aachen'
    response = requests.get(url)
    html_source = response.text
    return html_source


def find(xpath):
    """
    Returns the element specified by an XPATH, returns 0 if the element can not be found
    """""
    try:
        element = driver.find_element(By.XPATH, xpath)
        return element
    except Exception:
        return 0


def click(element):
    """
    Clicks on an element, will wait up to 5 seconds for the element to load
    """""
    wait = WebDriverWait(driver, 5)
    wait.until(EC.element_to_be_clickable(element)).click()


def text(element):
    """
    Returns the text of an element, replaces '/' with ''
    """""
    text = element.text
    if text == "KopernikusstraßeSeffenter Weg":
        new_text = "Kopernikusstraße Seffenter Weg"
    else:
        new_text = text.replace("/", "")
    return new_text


def bild(element):
    """
    Returns the name of the image 
    """""
    image_url = element.get_attribute('src')
    components = os.path.split(image_url)
    file_name = os.path.basename(components[-1])
    return file_name


def save(element, extracted_text):
    """
    Moves the mouse to the specified position and extracts the date
    """""
    pyautogui.moveTo(1880, 200)
    pyautogui.moveTo(1879, 200)
    controllDate = text(find(
        "/html/body/app-root/rit-dashboard/rit-dialog/div/div/div[2]/rit-sensor-things-widget/div/div[2]/div/div[2]"))
    date_regex = r"\d{2}\.\d{2}\."
    match = re.search(date_regex, controllDate)
    number = element.text
    if (len(number) > 0):
        lines = number.splitlines()
        del lines[0]
        if (match.group() == dateDE[:-4]):
            return (extracted_text, lines)
        else:
            if (len(lines) == 4):
                return (extracted_text, [lines[0], 0, lines[2], 0])
            elif (len(lines) == 2):
                return (extracted_text, (lines[0], 0))
    else:
        lines = 0
        return (extracted_text, lines)


def getDataDashboard(bool):
    """
    Navigates to the specified part of the website and gathers all the available data 
    """""
    click(find(
        "/html/body/app-root/rit-dashboard/div[1]/gridster/gridster-item[8]/rit-sensor-things-widget/div[1]/i[4]"))
    time.sleep(2)
    click(find(
        "/html/body/app-root/rit-dashboard/div[1]/gridster/gridster-item[8]/rit-sensor-things-widget/div[1]/i[2]"))
    time.sleep(1)
    click(
        find("/html/body/app-root/rit-dashboard/rit-dialog/div/div/div[1]/div/i[1]"))
    time.sleep(1)

    common_xpath = '/html/body/app-root/rit-dashboard/rit-dialog/div/div/div[2]/rit-sensor-things-widget/div/div[2]/div/div[2]'

    for i in range(1000):
        try:
            click(find(
                "/html/body/app-root/rit-dashboard/rit-dialog/div/div/div[2]/rit-sensor-things-widget/div/div[3]/ul/li["+str(i+1)+"]"))
            texts, lines = save(find(common_xpath), text(find(
                "/html/body/app-root/rit-dashboard/rit-dialog/div/div/div[2]/rit-sensor-things-widget/div/div[3]/ul/li["+str(i+1)+"]")))
            saveToExcel(texts, lines)
            csvParser(lines)
        except:
            break


def getWeatherData(content):
    """
    Gathers all the needed weather data, parses the data to fit the needs
    """""
    regexes = [r'{"date":"'+date+'","windSpeed":\d+,"temperatureMin":-?\d+,"temperatureMax":-?\d+,"precipitation":\d+}',
               r'{"date":"'+date+'","windSpeed":\d+,"temperatureMin":-?\d+,"temperatureMax":-?\d+,"precipitation":\d+\.\d+}',
               r'{"date":"'+date+'","windSpeed":\d+,"temperatureMin":-?\d+,"temperatureMax":-?\d+\.\d+,"precipitation":\d+}',
               r'{"date":"'+date+'","windSpeed":\d+,"temperatureMin":-?\d+,"temperatureMax":-?\d+\.\d+,"precipitation":\d+\.\d+}',
               r'{"date":"'+date+'","windSpeed":\d+,"temperatureMin":-?\d+\.\d+,"temperatureMax":-?\d+,"precipitation":\d+}',
               r'{"date":"'+date+'","windSpeed":\d+,"temperatureMin":-?\d+\.\d+,"temperatureMax":-?\d+,"precipitation":\d+\.\d+}',
               r'{"date":"'+date+'","windSpeed":\d+,"temperatureMin":-?\d+\.\d+,"temperatureMax":-?\d+\.\d+,"precipitation":\d+}',
               r'{"date":"'+date+'","windSpeed":\d+,"temperatureMin":-?\d+\.\d+,"temperatureMax":-?\d+\.\d+,"precipitation":\d+\.\d+}',
               r'{"date":"'+date+'","windSpeed":\d+\.\d+,"temperatureMin":-?\d+,"temperatureMax":-?\d+,"precipitation":\d+}',
               r'{"date":"'+date+'","windSpeed":\d+\.\d+,"temperatureMin":-?\d+,"temperatureMax":-?\d+,"precipitation":\d+\.\d+}',
               r'{"date":"'+date+'","windSpeed":\d+\.\d+,"temperatureMin":-?\d+,"temperatureMax":-?\d+\.\d+,"precipitation":\d+}',
               r'{"date":"'+date+'","windSpeed":\d+\.\d+,"temperatureMin":-?\d+,"temperatureMax":-?\d+\.\d+,"precipitation":\d+\.\d+}',
               r'{"date":"'+date+'","windSpeed":\d+\.\d+,"temperatureMin":-?\d+\.\d+,"temperatureMax":-?\d+,"precipitation":\d+}',
               r'{"date":"'+date+'","windSpeed":\d+\.\d+,"temperatureMin":-?\d+\.\d+,"temperatureMax":-?\d+,"precipitation":\d+\.\d+}',
               r'{"date":"'+date+'","windSpeed":\d+\.\d+,"temperatureMin":-?\d+\.\d+,"temperatureMax":-?\d+\.\d+,"precipitation":\d+}',
               r'{"date":"'+date+'","windSpeed":\d+\.\d+,"temperatureMin":-?\d+\.\d+,"temperatureMax":-?\d+\.\d+,"precipitation":\d+\.\d+}']
    for regex in regexes:
        matches = re.findall(regex, content)
        if len(matches) != 0:
            break
    for match in matches:
        substring = match
        substring = substring.replace('"', '')
        substring = substring.replace('{', '')
        substring = substring.replace('}', '')
        datastring = substring.replace('date:', '')
        datastring = datastring.replace('windSpeed:', '')
        datastring = datastring.replace('temperatureMin:', '')
        datastring = datastring.replace('temperatureMax:', '')
        datastring = datastring.replace('precipitation:', '')
        allvalues = substring.split(',')
        datavalues = datastring.split(',')
        pattern = re.compile(r"(\w+):(\d+\.\d+)")
        dataNames = ["Min. Tempearatur in °C",
                     "Max. Temperatur °C", "Niederschlagsmenge in l/m²"]
        del datavalues[0]
        del datavalues[0]
        for i, data in enumerate(datavalues):
            datavalues[i] = data.replace('.', ',')
        saveToExcel2(dataNames, datavalues)
        csvDataCollect(datavalues, 1)


def getBilder(content):
    """
    Gathers the name of the weather icon of today
    """""
    substringList = []
    trimmed_date = dateDEToday[:-4]
    regex = r'<td class="" data-tt-args="\[&quot;'+trimmed_date+'&quot;,&quot;.*&quot;,&quot;.*&quot;,.,0, 0, &quot;&quot;, &quot;&quot;,0, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;]" data-tt-function="TTwwsym">[\r\n]+ <img src="https:\/\/st\.wetteronline\.de\/dr\/1\.1\..*\/city\/prozess\/graphiken\/symbole\/standard\/farbe\/png\/[0-9][0-9]x[0-9][0-9]\/.*\.png'
    regexes = [r'bd____.png',
               r'bdg1__.png',
               r'bdg2__.png',
               r'bdgr1_.png',
               r'bdgr2_.png',
               r'bdr1__.png',
               r'bdr2__.png',
               r'bdr3__.png',
               r'bdsg__.png',
               r'bdsn1_.png',
               r'bdsn2_.png',
               r'bdsn3_.png',
               r'bdsr1_.png',
               r'bdsr2_.png',
               r'bdsr3_.png',
               r'nb____.png',
               r'ns____.png',
               r'so____.png',
               r'wb____.png',
               r'wbg1__.png',
               r'wbg2__.png',
               r'wbs1__.png',
               r'wbs2__.png',
               r'wbsg__.png',
               r'wbsns1.png',
               r'wbsns2.png',
               r'wbsrs1.png',
               r'wbsrs2.png',
               r'wbr1__.png',
               r'wbr2__.png',
               r'bw____.png',
               r'bws2__.png',
               r'bwr1__.png',
               r'bwr2__.png',
               r'bwg1__.png']
    matches = re.findall(regex, content)
    match = ', '.join(matches)
    for regex in regexes:
        substrings = re.findall(regex, match)
        if len(substrings) != 0:
            break
    substring = ', '.join(substrings)
    if (not substring):
        saveToExcel3(match)
        substringList.append(match)
    else:
        saveToExcel3(substring)
        substringList.append(substring)
    csvDataCollect(substringList, 0)


def startExcel():
    """
    Loads/Creates an excel file named data.xlsx
    """""
    if not os.path.exists('C:/Users/Fabian/Desktop/radwegzaehler/Excel/data.xlsx'):
        workbook = openpyxl.Workbook()
    else:
        workbook = openpyxl.load_workbook(
            'C:/Users/Fabian/Desktop/radwegzaehler/Excel/data.xlsx')
    return workbook


def counter2():
    """
    Increments the counter by two and returns it
    """""
    counter2.count += 2
    return counter2.count


def counter2minus1():
    """
    Decrements the counter by 1
    """""
    counter2.count -= 1


def date():
    """
    Gathers and returns the date from yesterday in format MM/DD/YYYY
    """""
    today = datetime.date.today()
    yesterday = today - timedelta(days=1)
    date_str = str(yesterday)
    return date_str


def dateDE():
    """
    Gathers and returns the date from yesterday in format DD/MM/YYYY
    """""
    today = datetime.date.today()
    yesterday = today - timedelta(days=1)
    date_str = yesterday.strftime("%d.%m.%Y")
    return date_str


def dateDEToday():
    """
    Gathers and returns the date from today in format DD/MM/YYYY
    """""
    today = datetime.date.today()
    date_str = today.strftime("%d.%m.%Y")
    return date_str


def findFirstEmptyCol():
    """
    Returns the first empty column that is not the first one
    """""
    if not os.path.exists('C:/Users/Fabian/Desktop/radwegzaehler/Excel/data.xlsx'):
        return 1
    else:
        path = "C:/Users/Fabian/Desktop/radwegzaehler/Excel/data.xlsx"
        df = pd.read_excel(path, sheet_name=0)
        empty_col = df.iloc[:, 1:].columns[(df.iloc[:, 1:].isna().all())]
        if empty_col.empty:
            number = df.shape[1]
            return number
        else:
            return empty_col[0]


def findLastSavedDate():
    """
    Finds the date from the last time the script ran, returns true if the last saved date is yesterdays date. Returns false if that is not the case
    """""
    if not os.path.exists('C:/Users/Fabian/Desktop/radwegzaehler/Excel/data.xlsx'):
        return False
    else:
        path = "C:/Users/Fabian/Desktop/radwegzaehler/Excel/data.xlsx"
        df = pd.read_excel(path)
        highest_column_value = df.iloc[0, df.shape[1] - 1]

        # Check if the value is equal to the variable 'dateDE'
        if highest_column_value == dateDE:
            return True
        else:
            return False


def saveToExcel(filenames, data):
    """
    Writes the data gathered about the zählstellen and yesterdays date into the excel file
    """""
    worksheet.cell(row=2, column=col_num+2).value = dateDE
    row_num = counter2()
    if type(data) == list:
        if len(data) == 2:
            worksheet.cell(row=row_num+4, column=col_num +
                           1).value = (filenames+" "+data[0])
            if (int(data[1]) != 0):
                worksheet.cell(
                    row=row_num+4, column=col_num+2).value = int(data[1])
            else:
                worksheet.cell(row=row_num+4, column=col_num +
                               2).value = "No Data"
            counter2minus1()
        if len(data) == 4:
            worksheet.cell(row=row_num+4, column=col_num +
                           1).value = (filenames+" "+data[0])
            worksheet.cell(row=row_num+4+1, column=col_num +
                           1).value = (filenames+" "+data[2])
            if (int(data[1]) != 0 and int(data[3]) != 0):
                worksheet.cell(row=row_num+4, column=col_num +
                               2).value = int(data[1])
                worksheet.cell(row=row_num+4+1, column=col_num +
                               2).value = int(data[3])
            else:
                worksheet.cell(row=row_num+4, column=col_num +
                               2).value = "No Data"
                worksheet.cell(row=row_num+4+1,
                               column=col_num+2).value = "No Data"
    if type(data) == int:
        worksheet.cell(row=row_num+4, column=col_num+1).value = (filenames)
        worksheet.cell(row=row_num+4, column=col_num+2).value = "No Data"
        if (filenames == "Vennbahntrasse (Ecke Philipsstraße)" or filenames == "Bismarckstraße" or filenames == "Königstraße" or filenames == "Lothringer Straße" or filenames == "Templergraben"):
            worksheet.cell(row=row_num+4+1, column=col_num+1).value = "No Data"
        else:
            counter2minus1()


def saveToExcel2(dataNames, datavalues):
    """
    Writes the weather data into the excel file
    """""
    worksheet.cell(row=4, column=col_num+1).value = dataNames[0]
    worksheet.cell(row=5, column=col_num+1).value = dataNames[1]
    worksheet.cell(row=6, column=col_num+1).value = dataNames[2]
    worksheet.cell(row=4, column=col_num+2).value = datavalues[0]
    worksheet.cell(row=5, column=col_num+2).value = datavalues[1]
    if datavalues[2] != "0":
        worksheet.cell(row=6, column=col_num+2).value = datavalues[2]


def saveToExcel3(imgName):
    """
    Writes the name of the weather icon to the excel file
    """""
    worksheet.cell(row=3, column=col_num+1).value = 'Wetter Symbol'
    if (not imgName):
        worksheet.cell(row=3, column=col_num+2).value = "No Data"
    else:
        worksheet.cell(row=3, column=col_num+2).value = imgName


def scale_column_width(column_number):
    """
    Scales the width of the column to 11, BROKEN
    """""
    column_letter = get_column_letter(column_number)
    column_dimensions = worksheet.column_dimensions
    column = column_dimensions[column_letter]
    column.width = 11


def checker():
    """
    Checks if all for the day data has been written into the excel file
    """""
    column_letter_1 = openpyxl.utils.get_column_letter(col_num+1)
    column_letter_2 = openpyxl.utils.get_column_letter(col_num + 2)
    column_index_1 = openpyxl.utils.column_index_from_string(column_letter_1)
    column_index_2 = openpyxl.utils.column_index_from_string(column_letter_2)
    column_1_empty = True
    for row in range(3, worksheet.max_row + 1):
        if (row != 6):
            cell_value1 = worksheet.cell(row=row, column=column_index_1).value
            if cell_value1 is None:
                column_1_empty = False
                print("Column 1 False")
    return column_1_empty


def csvDataCollect(data_list, index):
    """
    Collects all data into a list
    """""
    listDateDE = []
    listDateDE.append(dateDE)
    csvDataLists[0] = listDateDE
    if index == 0:
        csvDataLists[1] = data_list
    elif index == 1:
        csvDataLists[2] = data_list
    elif index == 2:
        csvDataLists[3] = data_list


def csvParser(lines):
    """
    Parses zählstellen data into one list
    """""
    if (isinstance(lines, int)):
        dataList.append(lines)
    elif (len(lines) == 4):
        dataList.append(lines[1])
        dataList.append(lines[3])
    elif (len(lines) == 2):
        dataList.append(lines[1])


def csvBackup():
    """
    Writes all data into a new csv file for each day, BROKEN
    """""
    x = 0
    text = []
    text.append("Bilder Name nicht passend zum Datum +1 Tag")
    csvDataLists.insert(0, text)
    csvFileName = "C:\\Users\\Fabian\\Desktop\\radwegzaehler\\Csv\\"+dateDE+".csv"
    if os.path.isfile(csvFileName):
        os.remove(csvFileName)
    with open(csvFileName, 'a', newline='') as file:
        writer = csv.writer(file)
        for csvDataList in csvDataLists:
            for data in csvDataList:
                x = x+1
                if (x == 3 and not data):
                    writer.writerow(["No Data"])
                else:
                    if (data != 0):
                        writer.writerow([data])
                    else:
                        writer.writerow(["No Data"])


# assign the value 1 to the count attribute of the counter2 object
counter2.count = 1
# create a date object representing the current date of yesterday
date = date()
# create a date object representing the current date in the German language of yesterday
dateDE = dateDE()
# create a date object representing the current date in the German language
dateDEToday = dateDEToday()
# find the first empty column in the worksheet and assign the result to the col_num variable
col_num = findFirstEmptyCol()
# create a new workbook object
workbook = startExcel()
# get the active worksheet in the workbook
worksheet = workbook.active
# 2d array for gathering data required for the csv backup
csvDataLists = [[], [], [], []]
# array for gathering data required for the csv backup
dataList = []

while True:
    try:
        # if the script has not run already today
        if findLastSavedDate() == False:
            while (True):
                counter2.count = 1
                driver = startDriver()
                # multithreading for the gathering and saving of the data from the three websites
                thread1 = threading.Thread(
                    target=lambda: getDataDashboard(startDashboard()))
                thread2 = threading.Thread(
                    target=lambda: getWeatherData(startWeather()))
                thread3 = threading.Thread(
                    target=lambda: getBilder(startPictures()))
                thread4 = threading.Thread(target=lambda: csvBackup())
                thread5 = threading.Thread(target=lambda: workbook.save(
                    "C:/Users/Fabian/Desktop/radwegzaehler/Excel/data.xlsx"))
                thread1.start()
                thread2.start()
                thread3.start()
                threads = [thread1, thread2, thread3]
                for thread in threads:
                    thread.join()
                thread4.start()
                thread5.start()
                driver.quit()
                thread4.join()
                thread5.join()
                scale_column_width(col_num)
                # if the excel file has been appropriately filled with data the loop breaks
                if (checker()):
                    break
        # if the scrip has ran already today the loop breaks
        else:
            break
    # if an exception occurs, print a message and continue the loop
    except:
        print("An error occurred, retrying...")
        continue
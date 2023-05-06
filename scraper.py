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
    # create a new Chrome webdriver
    newdriver = webdriver.Chrome()
    # return the new webdriver
    return newdriver


def startDashboard():
    # define the URL of the dashboard
    url = 'https://verkehr.aachen.de'
    # open the URL in the web browser
    driver.get(url)
    # maximize the browser window
    driver.maximize_window()
    # pause the script for 2 seconds
    time.sleep(2)
    return True


def startWeather():
    # define the URL of the weather website
    url = 'https://www.wetter.com/wetter_aktuell/rueckblick/deutschland/aachen/DE0000003.html?timeframe=30d'
    # send a GET request to the URL
    response = requests.get(url)
    # retrieve the HTML source code of the website
    html_source = response.text
    # return the HTML source code
    return html_source


def startBilder():
    # define the URL of the weather website
    url = 'https://www.wetteronline.de/wetter/aachen'
    # send a GET request to the URL
    response = requests.get(url)
    # retrieve the HTML source code of the website
    html_source = response.text
    # return the HTML source code
    return html_source


def find(xpath):
    try:
        element = driver.find_element(By.XPATH, xpath)
        # return the element
        return element
    except Exception:
        return 0
            

def click(element):
    # create a WebDriverWait object with a timeout of 5 seconds
    wait = WebDriverWait(driver, 5)
    # wait until the element is clickable, then click it
    wait.until(EC.element_to_be_clickable(element)).click()


def text(element):
    # get the text of the element
    text = element.text
    # check if the text needs to be modified
    if text == "KopernikusstraßeSeffenter Weg":
        new_text = "Kopernikusstraße Seffenter Weg"
    else:
        # replace slashes with an empty string
        new_text = text.replace("/", "")
    # return the modified text
    return new_text


def bild(element):
    # get the URL of the image
    image_url = element.get_attribute('src')
    # split the URL into components
    components = os.path.split(image_url)
    # get the file name from the URL
    file_name = os.path.basename(components[-1])
    # return the file name
    return file_name


def save(element, extracted_text):
    # Move the mouse cursor to the specified coordinates
    pyautogui.moveTo(1880, 200)
    pyautogui.moveTo(1879, 200)
    controllDate = text(find(
        "/html/body/app-root/rit-dashboard/rit-dialog/div/div/div[2]/rit-sensor-things-widget/div/div[2]/div/div[2]"))
    # matches two digits, a dot, two digits, and a dot
    date_regex = r"\d{2}\.\d{2}\."
    match = re.search(date_regex, controllDate)
    # Get the text of the element and assign it to the variable 'number'
    number = element.text
    if (len(number) > 0):
        lines = number.splitlines()
        del lines[0]
        if (match.group() == dateDE[:-4]):
            # Return a tuple containing 'text' and 'lines'
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
    # click on the icon to open the widget
    click(find(
        "/html/body/app-root/rit-dashboard/div[1]/gridster/gridster-item[8]/rit-sensor-things-widget/div[1]/i[4]"))
    time.sleep(2)
    # click on the icon to open the dialog
    click(find(
        "/html/body/app-root/rit-dashboard/div[1]/gridster/gridster-item[8]/rit-sensor-things-widget/div[1]/i[2]"))
    # close the dialog
    time.sleep(1)
    click(
        find("/html/body/app-root/rit-dashboard/rit-dialog/div/div/div[1]/div/i[1]"))
    time.sleep(1)

    common_xpath = '/html/body/app-root/rit-dashboard/rit-dialog/div/div/div[2]/rit-sensor-things-widget/div/div[2]/div/div[2]'
    
    for i in range(1000):
        try:
            click(find("/html/body/app-root/rit-dashboard/rit-dialog/div/div/div[2]/rit-sensor-things-widget/div/div[3]/ul/li["+str(i+1)+"]"))
            texts, lines = save(find(common_xpath), text(find("/html/body/app-root/rit-dashboard/rit-dialog/div/div/div[2]/rit-sensor-things-widget/div/div[3]/ul/li["+str(i+1)+"]")))
            saveToExcel(texts, lines)
            csvParser(lines)
        except:
            break
        

def getWeatherData(content):
    # list of regular expressions to match the weather data in the content
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
        # find all matches for the regex in the content
        matches = re.findall(regex, content)
        # break the loop if there are any matches
        if len(matches) != 0:
            break
    # loop through the matches
    for match in matches:
        # store the match in a substring
        substring = match
        # remove the quotes, curly braces, and colons from the substring
        substring = substring.replace('"', '')
        substring = substring.replace('{', '')
        substring = substring.replace('}', '')
        datastring = substring.replace('date:', '')
        datastring = datastring.replace('windSpeed:', '')
        datastring = datastring.replace('temperatureMin:', '')
        datastring = datastring.replace('temperatureMax:', '')
        datastring = datastring.replace('precipitation:', '')
        # split the substring into a list of values
        allvalues = substring.split(',')
        # split the datastring into a list of values
        datavalues = datastring.split(',')
        # compile a regex pattern to match key-value pairs
        pattern = re.compile(r"(\w+):(\d+\.\d+)")
        # define a list of data names
        dataNames = ["Min. Tempearatur in °C",
                     "Max. Temperatur °C", "Niederschlagsmenge in l/m²"]
        # remove the first two elements from the datavalues list
        del datavalues[0]
        del datavalues[0]
        # loop through the datavalues list
        for i, data in enumerate(datavalues):
            # replace dots with commas in the values
            datavalues[i] = data.replace('.', ',')
        # save the data to the Excel worksheet
        saveToExcel2(dataNames, datavalues)
        csvDataCollect(datavalues, 1)


def getBilder(content):
    substringList = []
    # get the current date and remove the time part
    trimmed_date = dateDEToday[:-4]
    # create a regular expression to match the desired content in the input
    regex = r'<td class="" data-tt-args="\[&quot;'+trimmed_date+'&quot;,&quot;.*&quot;,&quot;.*&quot;,.,0, 0, &quot;&quot;, &quot;&quot;,0, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;]" data-tt-function="TTwwsym">[\r\n]+ <img src="https:\/\/st\.wetteronline\.de\/dr\/1\.1\..*\/city\/prozess\/graphiken\/symbole\/standard\/farbe\/png\/[0-9][0-9]x[0-9][0-9]\/.*\.png'
    # create a list of regular expressions to match different substrings in the input
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
    # find all matches of the first regular expression in the input
    matches = re.findall(regex, content)
    # join the matches into a single string
    match = ', '.join(matches)
    # iterate over the list of regular expressions
    for regex in regexes:
        # find all matches of the current regular expression in the input
        substrings = re.findall(regex, match)
        # if at least one match is found, exit the loop
        if len(substrings) != 0:
            break
    # join the substrings into a single string
    substring = ', '.join(substrings)
    # save the substring to Excel
    if(not substring):
        saveToExcel3(match)
        substringList.append(match)
    else:
        saveToExcel3(substring)
        substringList.append(substring)
    csvDataCollect(substringList, 0)


def startExcel():
    # check if the specified Excel file exists
    if not os.path.exists('C:/Users/Fabian/Desktop/radwegzaehler/Excel/data.xlsx'):
        # create a new Excel workbook if the file does not exist
        workbook = openpyxl.Workbook()
    else:
        # load the existing Excel file if it exists
        workbook = openpyxl.load_workbook(
            'C:/Users/Fabian/Desktop/radwegzaehler/Excel/data.xlsx')
    # return the workbook object
    return workbook


def counter2():
    # increment the count by 2
    counter2.count += 2
    # return the updated count
    return counter2.count


def counter2minus1():
    # decrement the count by 1
    counter2.count -= 1


def date():
    # get the current date
    today = datetime.date.today()
    # calculate yesterday's date
    yesterday = today - timedelta(days=1)
    # convert the date to a string
    date_str = str(yesterday)
    # return the string
    return date_str


def dateDE():
    # get the current date
    today = datetime.date.today()
    # calculate yesterday's date
    yesterday = today - timedelta(days=1)
    # convert the date to a string in the dd.mm.yyyy format
    date_str = yesterday.strftime("%d.%m.%Y")
    # return the string
    return date_str


def dateDEToday():
    # get the current date
    today = datetime.date.today()
    # convert the date to a string in the dd.mm.yyyy format
    date_str = today.strftime("%d.%m.%Y")
    # return the string
    return date_str


def findFirstEmptyCol():
    # check if the Excel file exists
    if not os.path.exists('C:/Users/Fabian/Desktop/radwegzaehler/Excel/data.xlsx'):
        # return 2 if the file does not exist
        return 1
    else:
        # read the Excel file into a Pandas DataFrame
        path = "C:/Users/Fabian/Desktop/radwegzaehler/Excel/data.xlsx"
        df = pd.read_excel(path, sheet_name=0)
        # find the first empty column except the first column
        empty_col = df.iloc[:,1:].columns[(df.iloc[:,1:].isna().all())]
        # check if there are any empty columns
        if empty_col.empty:
            # return the number of columns if there are no empty columns
            number = df.shape[1]
            return number
        else:
            # return the first empty column except the first one
            return empty_col[0]




def findLastSavedDate():
    # Check if the file 'data.xlsx' exists in the specified path
    if not os.path.exists('C:/Users/Fabian/Desktop/radwegzaehler/Excel/data.xlsx'):
        return False
    else:
        # If the file exists, read it into a Pandas dataframe
        path = "C:/Users/Fabian/Desktop/radwegzaehler/Excel/data.xlsx"
        df = pd.read_excel(path)

        # Get the value in the second to last column of the first row of the dataframe
        highest_column_value = df.iloc[0, df.shape[1] - 1]

        # Check if the value is equal to the variable 'dateDE'
        if highest_column_value == dateDE:
            return True
        else:
            return False


def saveToExcel(filenames, data):
    # write the current date to cell B2
    worksheet.cell(row=2, column=col_num+2).value = dateDE
    row_num = counter2()
    # check the data type
    if type(data) == list:
        # if the data is a list with 2 elements
        if len(data) == 2:
            
            # write the filenames and data to the worksheet
            worksheet.cell(row=row_num+4, column=col_num+1).value = (filenames+" "+data[0])
            if (int(data[1]) != 0):
                worksheet.cell(
                    row=row_num+4, column=col_num+2).value = int(data[1])
            else:
                worksheet.cell(row=row_num+4, column=col_num+2).value = "No Data"
            # decrement the row number
            counter2minus1()
        # if the data is a list with 4 elements
        if len(data) == 4:
            # write the filenames and data to the worksheet
            worksheet.cell(row=row_num+4, column=col_num+1).value = (filenames+" "+data[0])
            worksheet.cell(row=row_num+4+1,column=col_num+1).value = (filenames+" "+data[2])
            if (int(data[1]) != 0 and int(data[3]) != 0):
                worksheet.cell(row=row_num+4, column=col_num+2).value = int(data[1])
                worksheet.cell(row=row_num+4+1,column=col_num+2).value = int(data[3])
            else:
                worksheet.cell(row=row_num+4, column=col_num+2).value = "No Data"
                worksheet.cell(row=row_num+4+1,column=col_num+2).value = "No Data"
    # if the data is an integer
    if type(data) == int:
        # write the filenames and "No Data" to the worksheet
        worksheet.cell(row=row_num+4, column=col_num+1).value = (filenames)
        worksheet.cell(row=row_num+4, column=col_num+2).value = "No Data"
        if(filenames=="Vennbahntrasse (Ecke Philipsstraße)" or filenames=="Bismarckstraße" or filenames=="Königstraße" or filenames=="Lothringer Straße" or filenames == "Templergraben"):
            worksheet.cell(row=row_num+4+1, column=col_num+1).value = "No Data"
        else:
            # decrement the row number
            counter2minus1()


def saveToExcel2(dataNames, datavalues):
    # write the data names and values to the worksheet
    worksheet.cell(row=4, column=col_num+1).value = dataNames[0]
    worksheet.cell(row=5, column=col_num+1).value = dataNames[1]
    worksheet.cell(row=6, column=col_num+1).value = dataNames[2]
    worksheet.cell(row=4, column=col_num+2).value = datavalues[0]
    worksheet.cell(row=5, column=col_num+2).value = datavalues[1]
    # check if the third data value is not "0"
    if datavalues[2] != "0":
        # write the third data value to the worksheet
        worksheet.cell(row=6, column=col_num+2).value = datavalues[2]


def saveToExcel3(imgName):
    # write the "Wetter Symbol" label to cell A3
    worksheet.cell(row=3, column=col_num+1).value = 'Wetter Symbol'
    if(not imgName):
        worksheet.cell(row=3, column=col_num+2).value = "No Data"
    else:
        # write the image name to cell C3
        worksheet.cell(row=3, column=col_num+2).value = imgName
        
    
def scale_column_width(column_number):
    # get the letter representation of the column number
    column_letter = get_column_letter(column_number)
    # get the column dimensions
    column_dimensions = worksheet.column_dimensions
    # get the column by its letter
    column = column_dimensions[column_letter]
    # set the column width to 11
    column.width = 11


def checker():
    # get the letter representation of the column numbers
    column_letter_1 = openpyxl.utils.get_column_letter(col_num+1)
    column_letter_2 = openpyxl.utils.get_column_letter(col_num + 2)
    # get the column indices from the letter representation
    column_index_1 = openpyxl.utils.column_index_from_string(column_letter_1)
    column_index_2 = openpyxl.utils.column_index_from_string(column_letter_2)
    # initialize the empty flags to True
    column_1_empty = True
    # loop through the rows in the worksheet
    for row in range(3, worksheet.max_row + 1):
        if (row != 6):
            # get the cell value in column 1
            cell_value1 = worksheet.cell(row=row, column=column_index_1).value
            # check if the cell value is None
            if cell_value1 is None:
                # set the empty flag to False
                column_1_empty = False
                print("Column 1 False")
    # return the result of the AND operation on the empty flags
    return column_1_empty


def csvDataCollect(data_list, index):
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
    if (isinstance(lines, int)):
        dataList.append(lines)
    elif (len(lines) == 4):
        dataList.append(lines[1])
        dataList.append(lines[3])
    elif (len(lines) == 2):
        dataList.append(lines[1])


def csvBackup():
    x = 0
    text = []
    text.append("Bilder Name nicht passend zum Datum +1 Tag")
    csvDataLists.insert(0, text)
    csvFileName = "C:\\Users\\Fabian\\Desktop\\radwegzaehler\\Csv\\"+dateDE+".csv"
    # check if file exists
    if os.path.isfile(csvFileName):
        # if file exists, delete it
        os.remove(csvFileName)
    # append data to file
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
# create a date object representing the current date
date = date()
# create a date object representing the current date in the German language
dateDE = dateDE()
# create a date object representing the current date in the German language, formatted as a string in the format "day.month.year"
dateDEToday = dateDEToday()
# find the first empty column in the worksheet and assign the result to the col_num variable
col_num = findFirstEmptyCol()
# create a new workbook object
workbook = startExcel()
# get the active worksheet in the workbook
worksheet = workbook.active
csvDataLists = [[], [], [], []]
dataList = []

while True:
    try:
        # check if the last saved date is false
        if findLastSavedDate() == False:
            while (True):
                # reset the counter2.count variable
                counter2.count = 1
                # start a new Chrome webdriver
                driver = startDriver()
                # create three threads to run the getDataDashboard, getWeatherData, and getBilder functions
                thread1 = threading.Thread(
                    target=lambda: getDataDashboard(startDashboard()))
                thread2 = threading.Thread(
                    target=lambda: getWeatherData(startWeather()))
                thread3 = threading.Thread(
                    target=lambda: getBilder(startBilder()))
                thread4 = threading.Thread(target=lambda: csvBackup())
                thread5 = threading.Thread(target=lambda: workbook.save(
                    "C:/Users/Fabian/Desktop/radwegzaehler/Excel/data.xlsx"))
                # start the threads
                thread1.start()
                thread2.start()
                thread3.start()
                # store the threads in a list
                threads = [thread1, thread2, thread3]
                # wait for all threads to complete
                for thread in threads:
                    thread.join()
                thread4.start()
                thread5.start()
                driver.quit()
                thread4.join()
                thread5.join()

                # scale the column width
                scale_column_width(col_num)
                # save the workbook
                # check if the checker function returns true
                if (checker()):
                    break
        # if the last saved date is not false, break the loop
        else:
            break
    # if an exception occurs, print a message and continue the loop
    except:
        print("An error occurred, retrying...")
        continue
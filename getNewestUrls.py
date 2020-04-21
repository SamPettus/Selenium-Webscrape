from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ChromeOptions
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import datetime
from bs4 import BeautifulSoup as bs
import re
import random
import xlwt
from xlwt import Workbook
import pandas as pd
EMAIL = #TO BE FILLED
PASSWORD = #TO BE FILLED
TARGET = 'Buisness'
TARGET2 = 'Technology'
fileUrl = 'linkedInUrls.xlsx'
class cardObject:
    def __init__(self, courseTitle, courseUrl):
        self.title = courseTitle
        self.url = courseUrl
def init_driver():
    opts = ChromeOptions()
    opts.add_experimental_option("detach", True)

    driver = webdriver.Chrome(options=opts)
    driver.wait = WebDriverWait(driver, 15)
    driver.maximize_window()
    return driver
def openUrlFile():
    df = pd.read_excel(fileUrl)
    return df
def login(driver):
    #login url
    driver.get('https://www.linkedin.com/learning/login?redirect=https%3A%2F%2Fwww.linkedin.com%2Flearning%2F%3Ftrk%3Ddefault_guest_learning&trk=sign_in')
    driver.wait.until(EC.url_contains(('https://www.linkedin.com/learning/login?redirect=https%3A%2F%2Fwww.linkedin.com%2Flearning%2F%3Ftrk%3Ddefault_guest_learning&trk=sign_in')))
    iframe = driver.find_element_by_tag_name('iframe')
    driver.switch_to.frame(iframe)
    #Types in username
    #usernameField = driver.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#auth-id-input')))
    usernameField = driver.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#username')))
    usernameField.send_keys(EMAIL)
    #usernameField..send_keys(u'\ue007')
    #Types in password
    passwordField = driver.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#password')))
    passwordField.send_keys(PASSWORD)
    loginButton = driver.find_element_by_xpath('//*[@id="app__container"]/main/div/form/div[3]/button')
    loginButton.click()
    driver.switch_to.default_content()

def searchByKeyword(driver, target):
    #Find the text field
    driver.wait.until(EC.url_contains(('https://www.linkedin.com/learning/me?trk=default_guest_learning')))
    driver.wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'ember-view')))
    content = driver.wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(concat(' ', @class, ' '), ' global-nav__content ')]/ul/li[3]/div/div/artdeco-typeahead/label/div/div/input")))
    #Searches target and hits enter
    content.send_keys(target)
    content.send_keys(u'\ue007')


def isolateCoursesAndNewest(driver):
    #UPDATE: LINKEDIN SAVES PRIOR SEARCH FIELD. NO LONGER NEED TO CLICK COURSES
    #Selects courses
    #xpathCourse = "/html/body/div[3]/div[1]/main/div/div/aside/div/div[1]/div/div/div/fieldset[1]/input"
    #coursesFeild = driver.wait.until(EC.presence_of_element_located((By.XPATH, xpathCourse)))
    #driver.execute_script('arguments[0].click();', coursesFeild)
    #time.sleep(2)
    #UPDATE OVER:
    #Clicks on drop down
    xpath = "//div[contains(concat(' ',@class, ' '), ' search-body__relevance-filter ')]/div/artdeco-dropdown/button"
    sortByDropDown = driver.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
    sortByDropDown.click()
    time.sleep(3)
    #Selects the newest tab
    xpath2 = "//div[contains(concat(' ',@class, ' '), ' search-body__relevance-filter ')]/div/artdeco-dropdown/artdeco-dropdown-content/ul/li[3]"
    newestField = driver.wait.until(EC.presence_of_element_located((By.XPATH, xpath2)))
    newestField.click()

#Checks to see if newest 100 elements have loaded
def checkForNewest1000(html):
    soup = bs(html, 'lxml')
    cards = soup.find_all('div', class_= 'lls-card-detail-card ember-view search-result-card')
    if len(cards) >= 1000:
        return True
    else:
        print(len(cards))
        return False

def infiniteScroll(driver, speed = 1, delay = .7, delayTime = 3):
    initialTime = time.time()
    count = 0
    current_scroll_position, new_height= 0, 1
    while current_scroll_position <= new_height:
        newTime = time.time()
        if count != 0:
            if new_height - current_scroll_position <=50:
                time.sleep(.5)
        if newTime - initialTime > delayTime:
            count = count + 1
            if count == 2:
                speed = 2
            if speed ==2:
                speed = 3
            temp = driver.page_source
            if checkForNewest1000(temp):
                return
            time.sleep(delay)
            delay = .3
            delayTime = 2
            initialTime = newTime
            if count % 50 == 0:
                speed = speed + 1
        current_scroll_position += speed
        driver.execute_script("window.scrollTo(0, {});".format(current_scroll_position))
        new_height = driver.execute_script("return document.body.scrollHeight")
    time.sleep(0.2)

def isolateCourseTitle(card):
    titleElement = card.find('h3', class_='card-layout-detail__title t-18 t-bold')
    if titleElement != None:
        return titleElement.text.strip()
    else:
        return 'NA'

def isolateUrl(card):
    urlElement = card.find('h3', 'card-layout-detail__title t-18 t-bold')
    if urlElement != None:
        href = urlElement.find('a', href=True)
        return 'https://www.linkedin.com' + href['href']
    else:
        return 'NA'

def createWorkBook():
    wb = Workbook()
    now = datetime.datetime.now()
    sheetName = '{}, {}, {}'.format(now.day, now.month, now.year)
    sheet1 = wb.add_sheet(sheetName)
    style = xlwt.easyxf('font: bold 1')
    #Creating Column Widths
    sheet1.col(0).width = 256 * 55
    sheet1.col(1).width = 256 * 65
    #Creating Column Headers
    sheet1.write(0, 0, 'Course Title', style)
    sheet1.write(0, 1, 'Url', style)
    return wb, sheet1

def createDataSet(html):
    soup = bs(html, 'lxml')
    cards = soup.find_all('div', class_= 'search-result-card card-layout-detail ember-view')
    listOfCards = []
    for card in cards:
        courseTitle = isolateCourseTitle(card)
        courseUrl = isolateUrl(card)
        obj = cardObject(courseTitle, courseUrl)
        listOfCards.append(obj)
    return listOfCards

def updateData(stored, dataBase):
    for i in stored:
        temp = dataBase.loc[dataBase['Url'] == i.url]
        if temp.empty:
            df1 = pd.DataFrame({'Course Title': [i.title], 'Url': [i.url]})
            dataBase = dataBase.append(df1, ignore_index=True)
        else:
            continue
    dataBase.drop_duplicates(subset= 'Url', keep = 'first')
    return dataBase
def find_nth(haystack, needle, n):
    start = haystack.find(needle)
    while start >= 0 and n > 1:
        start = haystack.find(needle, start+len(needle))
        n -= 1
    return start

def cleanData(data):
    for i in range(0,len(data)):
        substring = '/'
        count = data[i].url.count(substring)
        if count > 4:
            x = find_nth(data[i].url, substring, 5)
            obj = cardObject(data[i].title, data[i].url[:x])
            data[i] = obj
    removeDuplicate = []
    itemsToDelete = []
    #Find Duplicates
    for i in range(0, len(data)):
        if data[i].url not in removeDuplicate:
            removeDuplicate.append(data[i].url)
        else:
            itemsToDelete.append(i)
    #Delete Duplicates
    finalList = []
    for i in range(0, len(data)):
        if i not in itemsToDelete:
            finalList.append(data[i])

    return finalList

def main():
    #Previous runs
    dataBase = openUrlFile()
    driver = init_driver()
    #LogIn
    login(driver)
    #Buisness Search
    searchByKeyword(driver, TARGET)
    time.sleep(2)
    isolateCoursesAndNewest(driver)
    time.sleep(2)
    infiniteScroll(driver)
    time.sleep(1)
    html = driver.page_source
    #Second Search
    searchByKeyword(driver, TARGET2)
    isolateCoursesAndNewest(driver)
    time.sleep(2)
    infiniteScroll(driver)
    time.sleep(1)
    html2 = driver.page_source
    driver.quit()
    data1 = createDataSet(html)
    data2 = createDataSet(html2)
    combined = data1 + data2
    cleaned = cleanData(combined)
    df = updateData(cleaned, dataBase)
    df.to_excel('linkedInUrls.xlsx', index = None, header=True)

main()

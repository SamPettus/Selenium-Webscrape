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
EMAIL = #TO BE FILLED
PASSWORD = #TO BE FILLED
TARGET = 'Buisness'
TARGET2 = 'Technology'
class cardObject:
    def __init__(self, courseTitle, courseUrl):
        self.title = courseTitle
        self.url = courseUrl
def init_driver():
    opts = ChromeOptions()
    opts.add_experimental_option("detach", True)
    driver = webdriver.Chrome(chrome_options= opts)
    driver.wait = WebDriverWait(driver, 15)
    return driver

def login(driver):
    driver.get('https://www.linkedin.com/learning/login?redirect=https%3A%2F%2Fwww.linkedin.com%2Flearning%2F%3Ftrk%3Ddefault_guest_learning&trk=sign_in')
    driver.wait.until(EC.url_contains(('https://www.linkedin.com/learning/login?redirect=https%3A%2F%2Fwww.linkedin.com%2Flearning%2F%3Ftrk%3Ddefault_guest_learning&trk=sign_in')))
    iframe = driver.find_element_by_tag_name('iframe')
    driver.switch_to.frame(iframe)
    usernameField = driver.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#username')))
    usernameField.send_keys(EMAIL)
    passwordField = driver.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#password')))
    passwordField.send_keys(PASSWORD)
    loginButton = driver.find_element_by_xpath('//*[@id="app__container"]/main/div/form/div[3]/button')
    loginButton.click()
    driver.switch_to.default_content()

def searchByKeyword(driver, target):
    driver.wait.until(EC.url_contains(('https://www.linkedin.com/learning/me?trk=default_guest_learning')))
    driver.wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'ember-view')))
    content = driver.wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(concat(' ', @class, ' '), ' global-nav__content ')]/ul/li[3]/div/div/artdeco-typeahead/label/div/input")))
    content.send_keys(target + '\n')

def isolateCourses(driver):
    coursesFeild = driver.wait.until(EC.presence_of_element_located((By.ID, 'entityType-COURSE')))
    driver.execute_script('arguments[0].click();', coursesFeild)


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

def writeData(stored):
    wb, sheet1 = createWorkBook()
    courseNum = 1
    for i in stored:
        sheet1.write(courseNum, 0, i.title)
        sheet1.write(courseNum, 1, i.url)
        courseNum += 1
    now = datetime.datetime.now()
    workBookName = '{}({}, {}, {}).xls'.format('Urls', now.day, now.month, now.year)
    wb.save(workBookName)

def main():
    #Buisness Search
    driver = init_driver()
    login(driver)
    searchByKeyword(driver, TARGET)
    isolateCourses(driver)
    time.sleep(2)
    infiniteScroll(driver)
    time.sleep(1)
    html = driver.page_source
    driver.quit()
    #Technology Search
    driver = init_driver()
    login(driver)
    searchByKeyword(driver, TARGET2)
    isolateCourses(driver)
    time.sleep(2)
    infiniteScroll(driver)
    time.sleep(1)
    html2 = driver.page_source
    driver.quit()
    data1 = createDataSet(html)
    data2 = createDataSet(html2)
    combined = data1 + data2
    writeData(combined)

main()

import requests
from bs4 import BeautifulSoup as bs
import pandas as pd
from pandas import ExcelWriter
import time
import xlwt
from xlwt import Workbook
import xlsxwriter
import random
import datetime
from calendar import monthrange

fileUrl = 'linkedInUrls.xlsx'
#fileData = 'report.xls'
fileData = 'report.xlsx'
scrapedUrls = []
class cardObject:
    def __init__(self, courseTitle, courseViews, courseReleaseDate, coursePrice, courseTags):
        self.title = courseTitle
        self.views = courseViews
        self.price = coursePrice
        self.date = courseReleaseDate
        self.tags = courseTags
    def __lt__(self, other):
        return int(self.views.split()[0].replace(',','')) < int(other.views.split()[0].replace(',',''))
def openUrlFile():
    df = pd.read_excel(fileUrl)
    return df
def openDataFile():
    df = pd.read_excel(fileData)
    return df
def openDateSheet():
    df = pd.read_excel('report.xls', sheet_name='Days Scraped')
    return df
def getPrice(soup):
    priceElement = soup.find('button', class_='buy-course-upsell__cta buy-course-upsell__cta--buy-course')
    if priceElement != None:
        price = priceElement.text.split()[3]
        price = price[1:-2]
        return price.strip()
    else:
        return 'NA'
def getViews(soup):
    viewsElement = soup.find('span', class_='content__info__item__value viewers')
    if viewsElement != None:
        views = viewsElement.text
        return views.strip()
    else:
        return '0'
def getCourseName(soup):
    courseNameElement = soup.find('h1', class_='content__header-headline')
    if courseNameElement != None:
        courseName = courseNameElement.text
        return courseName.strip()
    else:
        return 'Skip'
def getReleaseDate(soup):
    releaseDateElement = soup.find('span', class_='content__info__item__value released')
    date = releaseDateElement.text
    return date.strip()
def getCourseTags(soup):
    tagsElement = soup.find('ul', class_='skills__list')
    if tagsElement != None:
        tags = tagsElement.text.split()
        return tags
    else:
        return 'NA'
def scrapeLinks(urlList):
    info = []
    count = 0
    for i in urlList:
        if (i in scrapedUrls):
            continue
        scrapedUrls.append(i)
        print(count)
        x = random.uniform(.3, 1.3)
        response = requests.get(i)
        soup = bs(response.text, 'lxml')
        title = getCourseName(soup)
        if title != 'Skip':
            price = getPrice(soup)
            views = getViews(soup)
            date = getReleaseDate(soup)
            tags = getCourseTags(soup)
            obj = cardObject(title, views, date, price, tags)
            info.append(obj)
        time.sleep(x)
        count += 1
    return info

def calculateNewSheet(information, previousRun):
    viewsSinceLastScrape = [0] * len(information)
    overlap = set(information['Course Title']).intersection(set(previousRun['Course Title']))
    information['weeklyViews'] = viewsSinceLastScrape
    for i in overlap:
        idx = information.index[information['Course Title'] == i]
        previousIdx = previousRun.index[previousRun['Course Title'] == i]
        prev = list(previousIdx.values)[0]
        curr = list(idx.values)[0]
        newViews = information.loc[curr, 'Views'].replace(',','')
        oldViews = previousRun.loc[prev, 'Views'].replace(',','')
        information.loc[idx, 'weeklyViews'] = int(newViews) - int(oldViews)
    return information

def convertListToDataFrame(information):
    dates = []
    titles = []
    prices = []
    views = []
    tags = []
    for i in information:
        dates.append(i.date)
        titles.append(i.title)
        tags.append(i.tags)
        views.append(i.views)
        prices.append(i.price)
    data = {'Course Title': titles, 'Tags': tags, 'Price': prices, 'Release Date': dates, 'Views': views}
    df = pd.DataFrame(data)
    return df

def writeData(df):
    workbook = xlsxwriter.Workbook('reportTEMP.xlsx')
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    cell_formatGain = workbook.add_format()
    cell_formatGain.set_font_color('green')
    cell_formatLoss = workbook.add_format()
    cell_formatLoss.set_font_color('red')
    worksheet.set_column('A:A', 50)
    worksheet.set_column('B:B', 70)
    worksheet.set_column('D:D', 50)

    worksheet.write('A1', 'Course Title', bold)
    worksheet.write('B1', 'Tags', bold)
    worksheet.write('C1', 'Price', bold)
    worksheet.write('D1', 'Release Date', bold)
    worksheet.write('E1', 'Views', bold)
    worksheet.write('F1', 'weeklyViews', bold)

    count = 1
    for index, row in df.iterrows():
        worksheet.write(count, 0, row['Course Title'])
        temp = ''
        for i in row['Tags']:
            temp = temp + i + ' '
        worksheet.write(count, 1, temp)
        worksheet.write(count, 2, row['Price'])
        worksheet.write(count, 3, row['Release Date'])
        worksheet.write(count, 4, row['Views'])
        if row['weeklyViews'] > 0:
            worksheet.write(count, 5, row['weeklyViews'], cell_formatGain)
        else:
            worksheet.write(count, 5, row['weeklyViews'], cell_formatLoss)
        count += 1
    workbook.close()


def main():
    #Opens file containing urls to scrape
    dataBaseUrl = openUrlFile()
    dataBaseList = dataBaseUrl['Url'].tolist()
    information = scrapeLinks(dataBaseList)
    #Opens previous data run
    previousRun = openDataFile()
    #dates = openDateSheet()
    df = convertListToDataFrame(information)
    final = calculateNewSheet(df, previousRun)
    final = final.sort_values(by=['weeklyViews'], ascending=False)
    writeData(final)


main()
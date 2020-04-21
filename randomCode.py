#Code to calculate days since last scraped
    #Math since last scraped
    #x = dates.iloc[-1]['Days Scraped']]
    #dt = datetime.datetime.today()
    #dateList = x.split()
    #daysSinceLastScrape = 0
    #if int(dateList[1]) == dt.month:
    #    daysSinceLastScrape = dt.day - int(dateList[0])
    #elif int(dateList[1]) == dt.month - 1:
    #    monthrange(int(dateList[2]), int(dateList[3]))
    #    temp = monthrange[1] - int(dateList[0])
    #    daysSinceLastScrape = temp + dt.day
#Code to create original workbook
def createWorkBook():
    wb = Workbook()
    sheetName = 'Sheet 1'
    sheet1 = wb.add_sheet(sheetName)
    sheet2 = wb.add_sheet('Days Scraped')
    style = xlwt.easyxf('font: bold 1')
    #Creating Column Widths
    sheet1.col(0).width = 256 * 55
    sheet1.col(1).width = 256 * 70
    sheet1.col(2).width = 256 * 20
    sheet1.col(3).width = 256 * 20
    sheet1.col(4).width = 256 * 20
    #Creating Column Headers
    sheet1.write(0, 0, 'Course Title', style)
    sheet1.write(0, 1, 'Tags', style)
    sheet1.write(0, 2, 'Price', style)
    sheet1.write(0, 3, 'Release Date', style)
    sheet1.write(0, 4, 'Views', style)

    sheet2.write(0,0, 'Days Scraped', style)
    dt = datetime.datetime.today()
    day = str(dt.day) + ' ' + str(dt.month) + ' ' + str(dt.year)
    sheet2.write(1, 0, day)
    return wb, sheet1, sheet2

def writeData(stored):
    wb, sheet1, sheet2 = createWorkBook()
    courseNum = 1
    for i in stored:
        sheet1.write(courseNum, 0, i.title)
        sheet1.write(courseNum, 1, i.tags)
        sheet1.write(courseNum, 2, i.price)
        sheet1.write(courseNum, 3, i.date)
        sheet1.write(courseNum, 4, i.views)
        courseNum += 1
    workBookName = 'report.xls'
    wb.save(workBookName)

# coding=utf-8
from selenium.webdriver import Firefox
from selenium.webdriver.support.events import EventFiringWebDriver, AbstractEventListener
import xlrd
import xlwt
from xlutils.copy import copy
import time
from selenium.webdriver.support.ui import WebDriverWait
import pickle


book_name_xls = 'xls格式测试工作簿.xls'


def write_excel_xls(path, sheet_name, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlwt.Workbook()  # 新建一个工作簿
    sheet = workbook.add_sheet(sheet_name)  # 在工作簿中新建一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            sheet.write(i, j, value[i][j])  # 像表格中写入数据（对应的行和列）
    workbook.save(path)  # 保存工作簿
    print("xls格式表格写入数据成功！")


def write_excel_xls_append(path, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            new_worksheet.write(i + rows_old, j, value[i][j])  # 追加写入数据，注意是从i+rows_old行开始写入
    new_workbook.save(path)  # 保存工作簿
    print("xls格式表格【追加】写入数据成功！")


def read_excel_xls(path):
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    for i in range(0, worksheet.nrows):
        for j in range(0, worksheet.ncols):
            print(worksheet.cell_value(i, j), "\t", end="")  # 逐行逐列读取数据
        print()


book_name_xls = '微博爬虫.xls'

sheet_name_xls = 'xls格式测试表'

value_title = [["用户链接", "用户微博名称", "微博内容","VIP身份","时间", "来源"], ]




write_excel_xls(book_name_xls, sheet_name_xls, value_title)




class MyListener(AbstractEventListener):
    def before_navigate_to(self, url, driver):
        print("Before navigate to %s" % url)
    def after_navigate_to(self, url, driver):
        print("After navigate to %s" % url)


driver = Firefox()
wait = WebDriverWait(driver, 2)
ef_driver=EventFiringWebDriver(driver,MyListener())
ef_driver.get("https://weibo.com")
for j in range(30):
    print(j)
    time.sleep(1)

ef_driver.get("https://s.weibo.com/weibo?q=ask%20me%20anything&nodup=1")

def nextPage():
    for h in range(5):
        print(h)
        time.sleep(1)
    contentBody = driver.find_element_by_id("pl_feedlist_index")
    nrows = contentBody.find_elements_by_class_name("card-wrap")
    for i in range(len(nrows)):
        print (i)
        avator = nrows[i].find_elements_by_class_name("avator")[0]
        # img = avator.find_elements_by_xpath('img')[0].get_attribute('src')
        # print (img)
        userHref = nrows[i].find_elements_by_class_name("name")[0].get_attribute("href")
        print (userHref)
        userName = nrows[i].find_elements_by_class_name("name")[0].text
        print (userName)
        content = nrows[i].find_elements_by_class_name("txt")[0].text
        print (content)
        vipList = nrows[i].find_elements_by_class_name("info")[0].find_elements_by_class_name("icon-vip")
        vip = ""
        if (len(vipList) > 0):
            vip = vipList[0].get_attribute("class")
            print (vip)
        fromList = nrows[i].find_elements_by_class_name("from")[0].find_elements_by_xpath('a')
        fromTime = fromList[0].text
        print (fromTime)
        fromPhone = ""
        if(len(fromList)>1):
            fromPhone = fromList[1].text
            print (fromPhone)
        print ("\n")
        value1 = [[userHref, userName, content, vip, fromTime, fromPhone]]
        write_excel_xls_append(book_name_xls, value1)
    time.sleep(3)
    driver.find_elements_by_class_name("m-page")[0].find_elements_by_class_name("next")[0].click()
    nextPage()

nextPage()
# icon-vip icon-vip-y  黄V
# icon-vip icon-member  会员
# icon-vip icon-daren  达人
# icon-vip icon-vip-b  蓝V




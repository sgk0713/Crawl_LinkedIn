# encoding=utf8

# python 3.6

import csv
import time
import logger

import openpyxl
from bs4 import BeautifulSoup as bs
from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys

import config_constants

Login_Button_XPATH = '//*[@id="main-content"]/section[1]/div[2]/form/button'
SEARCH_FIELD_XPATH = '//*[@id="mn-connections-search-input"]'
SEARCH_RESULT_FIRST_ITEM_SELECTOR = 'span.mn-connection-card__name.t-16.t-black.t-bold'

NAME_FIELD_SELECTOR = 'div.display-flex.justify-space-between.pt2 > div > div:nth-child(1) > h1'
CONTACT_FIELD_SELECTOR = 'div.pb2 > span.pv-text-details__separator.t-black--light > a'

WEBSITE_ROOT_IN_CONTACT_SELECTOR = 'section.pv-contact-info__contact-type.ci-websites > ul li'
EMAIL_FIELD_IN_CONTACT_SELECTOR = 'div > section.pv-contact-info__contact-type.ci-email > div > a'

names = []  # put ur 1st connection's names(for now...will fix to use easily soon)

CONST_ID = config_constants.CONST_ID  # read from config_constants.py. put your infos at config_constants.py file
CONST_PW = config_constants.CONST_PW
DRIVER_PATH = config_constants.CONST_DRIVER_PATH


def do():
    driver = webdriver.Chrome(DRIVER_PATH)
    driver.implicitly_wait(1)
    # login
    driver.get('https://www.linkedin.com/')
    driver.find_element_by_name('session_key').send_keys(CONST_ID)
    driver.find_element_by_name('session_password').send_keys(CONST_PW)
    driver.find_element_by_xpath(Login_Button_XPATH).click()

    # get length of list
    max = len(names)
    main = driver.get('https://www.linkedin.com/mynetwork/invite-connect/connections/')
    html = driver.page_source
    soup = bs(html, 'html.parser')

    f = open('linkedin_list.csv', 'w', encoding='utf-8', newline='')
    wr = csv.writer(f)
    wr.writerow(['Number', 'Name', 'Email', 'HomePage'])

    errorFile = open('error.csv', 'w', encoding='utf-8', newline='')
    ewr = csv.writer(errorFile)
    ewr.writerow(['i'])

    for i in range(0, max):
        # na = unicode(name, "utf8", errors="ignore")

        if (True):
            try:
                driver.find_element_by_xpath(SEARCH_FIELD_XPATH).clear()
                driver.find_element_by_xpath(SEARCH_FIELD_XPATH).clear()
                driver.find_element_by_xpath(SEARCH_FIELD_XPATH).send_keys(names[i])
                time.sleep(2)
                element = driver.find_element_by_css_selector(SEARCH_RESULT_FIRST_ITEM_SELECTOR)

                # ActionChains(driver).key_down(Keys.SHIFT).click(element).key_up(Keys.SHIFT).perform()

                ActionChains(driver).key_down(Keys.COMMAND).click(element).key_up(Keys.COMMAND).perform()
                # time.sleep(2)
                driver.switch_to.window(driver.window_handles[1])
                name = driver.find_element_by_css_selector(
                    NAME_FIELD_SELECTOR).text
                temp = driver.find_element_by_css_selector(
                    CONTACT_FIELD_SELECTOR)
                ActionChains(driver).click(temp).perform()
            except:
                driver.switch_to.window(driver.window_handles[0])
                ewr.writerow(['i == %d' % i])
                continue
            homepages = []
            email = ""
            try:
                lists = driver.find_elements_by_css_selector(
                    WEBSITE_ROOT_IN_CONTACT_SELECTOR)
                for list in lists:
                    homepages.append(list.find_element_by_css_selector('div a').get_attribute("href"))
            except:
                homepages = None
            try:
                email = driver.find_element_by_css_selector(
                    EMAIL_FIELD_IN_CONTACT_SELECTOR).text
            except:
                email = None
            try:
                driver.close();
                driver.switch_to.window(driver.window_handles[0])
                print("#", i, ' >> ', email)
                list = []
                list.append(i)
                list.append(name)
                list.append(email)
                if (homepages is not None):
                    for homepage in homepages:
                        list.append(homepage)
                wr.writerow(list)
            except:
                driver.switch_to.window(driver.window_handles[0])
                ewr.writerow(['i == %d' % i])
                continue
        # if(i==5):
        # 	break
    f.close()
    errorFile.close()


def checkFile():
    wb = openpyxl.load_workbook('linkedIn_Final.xlsx')
    real = openpyxl.load_workbook('linkedin_list.xlsx')

    ws = wb.active
    re = real.active

    errorFile = open('errorCheck.csv', 'w', encoding='utf-8', newline='')
    ewr = csv.writer(errorFile)

    max = 3247

    wt = Workbook()

    # 파일 이름을 정하고, 데이터를 넣을 시트를 활성화합니다.
    sheet1 = wt.active
    file_name = 'test.xlsx'

    # 시트의 이름을 정합니다.
    sheet1.title = 'Sheet1'

    # cell 함수를 이용해 넣을 데이터의 행렬 위치를 지정해줍니다.
    for row_index in range(2, max):
        sheet1.cell(row=row_index, column=1).value = ws.cell(row_index, 1).value
        sheet1.cell(row=row_index, column=2).value = ws.cell(row_index, 2).value
        sheet1.cell(row=row_index, column=3).value = re.cell(row_index, 10).value
        sheet1.cell(row=row_index, column=4).value = ws.cell(row_index, 3).value
        sheet1.cell(row=row_index, column=5).value = ws.cell(row_index, 4).value
        sheet1.cell(row=row_index, column=6).value = ws.cell(row_index, 5).value
        sheet1.cell(row=row_index, column=7).value = ws.cell(row_index, 6).value

    for i in range(2, max - 1):

        for j in range(i + 1, max):
            first = ws.cell(i, 3).value
            second = ws.cell(j, 3).value
            if (first != None and second != None and first == second):
                ewr.writerow([i - 2, j - 2, first, second])
                sheet1.cell(j, 4).value = ""
                logger.log(f'{i}, {j}, {first}, {second}')
                continue

    wb.close()
    wt.save(filename=file_name)
    errorFile.close()


checkFile()
# do()

import xlsxwriter
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

import sys

curpath = sys.path[0]
print(curpath)

from bs4 import BeautifulSoup

DOWNLOAD_URL = 'https://yz.chsi.com.cn/zsml/querySchAction.do?ssdm=&dwmc=&mldm=zyxw&mlmc=&yjxkdm=0852&xxfs=1&zymc=%E8%BD%AF%E4%BB%B6%E5%B7%A5%E7%A8%8B####'

school_link_list = []


def get_html(html):
    soup = BeautifulSoup(html, 'html.parser')
    school_list_soup = soup.find('table', attrs={'class': 'ch-table'}).find('tbody')

    for school_li in school_list_soup.find_all('tr'):
        detail = school_li.find_all('td', attrs={'class': 'ch-table-center'})
        for find_school_link in detail:
            if find_school_link(text='查看'):
                school_link = find_school_link.find('a').get('href')
                school_link = "https://yz.chsi.com.cn" + school_link
                # print(school_link)
                school_link_list.append(school_link)
    return school_link_list


def get_data(school_link, driver):
    html = download_page(school_link, driver)
    soup = BeautifulSoup(html, 'html.parser')

    subject_info_list = list()
    school_info = soup.find('table', attrs={'class': 'zsml-condition'}).find('tbody')
    title_list = list()
    summary_list = list()
    for school_info_li in school_info.find_all('tr'):
        for school_info_title_li in school_info_li.find_all('td', attrs={'class': 'zsml-title'}):
            title = school_info_title_li.getText()
            title_list.append(title)
        for school_info_summary_li in school_info_li.find_all('td', attrs={'class': 'zsml-summary'}):
            summary = school_info_summary_li.getText()
            summary_list.append(summary)
    school_info_dict = dict(zip(title_list, summary_list))
    # print(school_info_dict)

    test_range_info = soup.find('div', attrs={'class': 'zsml-result'}).find('table')
    for test_range_info_li in test_range_info.find_all('tbody', attrs={'class': 'zsml-res-items'}):
        for subject_info_li in test_range_info_li.find('tr').find_all('td'):
            subject_info_list.append(subject_info_li.getText().strip('\r\n                '))
    # print(subject_info_list)
    return school_info_dict, subject_info_list


def download_page(url, driver):
    driver.get(url)
    data = driver.page_source
    return data


def main():
    url = DOWNLOAD_URL

    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    driver = webdriver.Chrome('D:/chromedriver.exe')
    driver.set_page_load_timeout(30)
    time.sleep(2)

    workbook = xlsxwriter.Workbook('data.xlsx')
    worksheet = workbook.add_worksheet()
    row = 0
    html = download_page(url, driver)
    school_link_list = get_html(html)
    for i in range(2, 20):
        searchInput = driver.find_element_by_id("goPageNo")
        # searchInput = driver.find_element_by_class_name("page-input")
        searchInput.send_keys(i)
        searchSubmitBtn = driver.find_element_by_class_name("page-btn")
        searchSubmitBtn.click()

        html = driver.page_source
        get_html(html)
    # print(school_link_list)
    for school_link in school_link_list:
        col = 0
        school_info_dict, subject_info_list = get_data(school_link, driver)
        for school_info_dict_li in school_info_dict.values():
            worksheet.write(row, col, school_info_dict_li)
            col += 1
        for subject_info_li in subject_info_list:
            worksheet.write(row, col, subject_info_li)
            col += 1
        row += 1
    workbook.close()
    # for school_link in school_link_list:
    #     school_info_dict, subject_info_list = get_data(school_link,driver)
    #     print(school_info_dict, subject_info_list)


if __name__ == '__main__':
    main()

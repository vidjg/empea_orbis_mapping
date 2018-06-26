# -*- coding: utf-8 -*-
"""
Created on Mon Jun 25 08:57:10 2018

@author: sqian
"""


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup as soup
from bs4 import SoupStrainer
from lxml import html
import numpy as np
import time
import pandas as pd
import win32com.client as win32
import os


def login_orbis(browser):
    # Turn on the broswer
    browser.get(login_url)
    # Login Page
    username = browser.find_element_by_name("user")
    password = browser.find_element_by_name("pw")
    username.send_keys(form_data["user"])
    password.send_keys(form_data["pw"])
    login_button = browser.find_element_by_class_name("ok")
    login_button.click()
    try:
        restart_button = browser.find_element_by_xpath("//input[@class='button ok']")
        restart_button.click()
    except:
        pass
                                                    

def hard_refresh(browser,year,start_page):
    browser.close()
    
    browser = webdriver.Chrome()
    browser.get(login_url)
    username = browser.find_element_by_name("user")
    password = browser.find_element_by_name("pw")
    username.send_keys(form_data["user"])
    password.send_keys(form_data["pw"])
    while 1:
        try:
            login_button = browser.find_element_by_class_name("ok")
            login_button.click()
            break
        except:
            browser.close()
            browser = webdriver.Chrome()
            browser.get(login_url)
    
    login_orbis(browser,year)
    while 1:
        try:
            page_input = browser.find_elements_by_css_selector("ul.navigation > li > input")[0]
            page_input.clear()
            page_input.send_keys(str(start_page))
            page_input.send_keys(Keys.RETURN)
            break
        except:
            continue
    refresh_stuck = visible_in_time(browser, '#resultsTable > tbody > tr > td.scroll-data > div > table > tbody > tr:nth-child(1) > td:nth-child(1)', 20)
    if refresh_stuck == False:
        return hard_refresh(browser,year,start_page)
    else:
        return browser
 
    
def visible_in_time(browser, address, time):
    try:
        WebDriverWait(browser, time).until(EC.presence_of_element_located((By.CSS_SELECTOR, address)))
        return True
    except TimeoutException:
        return False


def split_table(file_name):
    datatable = pd.read_csv(file_name)
    for i in range(len(datatable)//1000+1):
        datatable.loc[i*1000:(i+1)*1000].to_csv(file_name[:-4] + '_' + str(i) + '.csv', mode='w', index=True, sep='\t')


def select_score(browser, total_page_num, start_page):
    # Go to Page 1
    page_input = browser.find_element_by_xpath('//li/input[@type="number"]')
    page_input.clear()
    page_input.send_keys(start_page)
    page_input.send_keys(Keys.RETURN)
    print("Start Mapping!")
    time.sleep(1)
    for page in range(start_page-1,total_page_num):
        innerHTML = browser.execute_script("return document.body.innerHTML")
        tree = html.fromstring(innerHTML)
        orbis_result = tree.xpath('//*[@id="matchedSelected"]/@data-selected')
        orbis_content = tree.xpath('//*[@id="matchedSelected"]/@data-matched')
        for id in range(len(orbis_result)):
            if orbis_result[id] == '[]' and orbis_content[id] != '[]':
                nation = browser.find_element_by_xpath('//tbody/tr[@data-id={0}]/td[@data-id="Country"]/div'.format(id+page*100)).text
                item_expand = browser.find_element_by_xpath('//tbody/tr[@data-id={0}]/td/label'.format(id+page*100))
                item_expand.click()
                try:
                    score = browser.find_element_by_xpath('//*[@id="Template" and not(string(@style))]/td[7]/div').text
                    nation_popped = browser.find_element_by_xpath('//*[@id="Template" and not(string(@style))]/td[4]/div').text
                    if nation == nation_popped or len(nation) == 3:
                        browser.find_element_by_xpath('//*[@id="Template" and not(string(@style))]/td[1]/label').click()
                        print(id+page*100)
                        time.sleep(1.5)
                    else:
                        print("Mapping result rejected!")
                except:
                    continue
        browser.find_element_by_xpath('/html/body/div[2]/div[1]/div[4]/div/form/div[2]/ul/li[12]/img').click()


def create_mapping(browser, total_page_num):
    # Generate Mapping Table
    page_input = browser.find_element_by_xpath('//li/input[@type="number"]')
    page_input.clear()
    page_input.send_keys(1)
    page_input.send_keys(Keys.RETURN)
    time.sleep(1)
    print("Start Generating Mapping Table!")
    
    company_mapping = pd.DataFrame(columns=['original_name','country_code','mapped_name','mapped_bvdId','mapping_score'])
    
    for page in range(0,total_page_num):
        innerHTML = browser.execute_script("return document.body.innerHTML")
        tree = html.fromstring(innerHTML)
        data_temp = pd.DataFrame()
        data_temp['original_name'] = tree.xpath('//td[@data-id="Name"]/div/text()')
        data_temp['country_code'] = [x.text for x in tree.xpath('//td[@data-id="Country"]/div')]
        data_temp['mapped_name'] = [x.strip() for x in tree.xpath('//td[@id="matchedSelected"]/div/text()')]
        result_list = [eval(x) for x in tree.xpath('//td[@id="matchedSelected"]/@data-selected')]
        id_list = []
        for x in result_list:
            try:
                id_list.append(x[0]['BvDId'])
            except:
                id_list.append("")
        data_temp['mapped_bvdId'] = id_list
        data_temp['mapping_score'] = tree.xpath('//td[@id="matchedScore"]/div/@class')
        company_mapping = pd.concat([company_mapping,data_temp])
        browser.find_element_by_xpath('/html/body/div[2]/div[1]/div[4]/div/form/div[2]/ul/li[12]/img').click()
    
    # Output mapping table
    company_mapping.to_csv('mapping.csv', mode='a', index=False)
    return company_mapping


def data_scraping(browser):
    # Download Detailed Data
    select_view = browser.find_element_by_css_selector('div.menuViewContainer > div.menuView > ul > li > a')
    select_view.click()  
    if visible_in_time(browser, 'span.name.clickable[title="EMPEA"]', 30):
        view_year = browser.find_element_by_css_selector('span.name.clickable[title="EMPEA"]')
        view_year.click()
    
    innerHTML = browser.execute_script("return document.body.innerHTML")
    page_soup = soup(innerHTML, "html.parser")
    columns = page_soup.find('table', {'class': 'scroll-header'}).find_next('tr')
    label_info = columns.find_all('div', class_='column-label')
    column_names = []
    for x in label_info:
        try:
            column_names.append(x.span['data-fulllabel'] + ' <' + x.find_all('span')[1]['data-full-configuration'] + '>')
        except:
            column_names.append(x.span['data-fulllabel'])
    company_data = pd.DataFrame()
    company_names = []
    for x in column_names:
        company_data[x] = []
    
    per_page = 100
    total_companies = int(page_soup.find('td',{'class':'grand-total'}).text.replace(',',''))
    total_pages =  total_companies // per_page + 1 # Number of pages of data to retrieve
    page_done = 0
    
# =============================================================================
#     page_input = browser.find_elements_by_css_selector("ul.navigation > li > input")[0]
#     page_input.clear()
#     page_input.send_keys(str(1))
#     page_input.send_keys(Keys.RETURN)
# =============================================================================
    if visible_in_time(browser,'#resultsTable > tbody > tr > td.scroll-data > div > table > tbody > tr:nth-child(1) > td:nth-child(1)',20):
        pass
    else:
        print('Timeout!')
        exit()
    time.sleep(4)
    
    while page_done < total_pages:
        innerHTML = browser.execute_script("return document.body.innerHTML")
        tree = html.fromstring(innerHTML)
        page_num = int(tree.xpath('//ul[@class="navigation"]/*/span[@class="currentPage" and text() != "..." ]/text()')[0])
        if page_num == page_done+1 and visible_in_time(browser,'#resultsTable > tbody > tr > td.scroll-data > div > table > tbody > tr:nth-child(1) > td:nth-child(1)',20):
            print("Page {0} retrieved!".format(page_num))
            tree = tree.cssselect('#resultsTable')[0]
            if company_names == []:
                company_names = [x.text for x in tree.xpath('//span[@class="ellipsis"]/a[@href="#"]')]
            else:
                company_names += [x.text for x in tree.xpath('//span[@class="ellipsis"]/a[@href="#"]')]
            data_points = tree.xpath('//td[@class="scroll-data"]/div/table/tbody/tr/descendant::*/text()')
            if page_num == total_pages:
                num_on_page = len(tree.xpath('//span[@class="ellipsis"]/a[@href="#"]'))
            else:
                num_on_page = per_page
            data = np.array_split(data_points, num_on_page)                
            company_data = pd.concat([company_data,pd.DataFrame(data,columns=column_names)])
            
            page_done += 1
            print("Page {0} finished!".format(page_num))
            if page_num != total_pages:
                browser.find_element_by_xpath("//img[@data-action='next']").click()
    
    company_data.insert(0,"company_name",company_names)
    return company_data


def select_file(browser, file_name, file_id):
    # Turn to the page of Tools
    batch_search_page = "https://orbis4.bvdinfo.com/version-2018621/orbis/1/Companies/BatchSearch/Start"
    browser.get(batch_search_page)
    time.sleep(2)
    browser.find_element_by_id('upload-now').click()
    browser.find_element_by_css_selector('body > div.viewport.main > div.website > div.content > div > div.batchWidget > div > div > form > div.view > div:nth-child(1) > input.hidden').send_keys(os.getcwd()+'/' + file_name + '_{0}.csv'.format(file_id))
    browser.find_element_by_css_selector('dl.mapping-options > dd:nth-child(3) > label').click()
    browser.find_element_by_css_selector('div.batchWidget > div > div > form > div.buttons > div > a.button.ok').click()
    if visible_in_time(browser,'#CountDown',20):                  
        while 1:
            search_process = browser.find_element_by_css_selector('#CountDown').text.split('/')
            if search_process[0] == search_process[1]:
                time.sleep(5)
                break
            else:
                time.sleep(1)
                                                             

###### Main Function ######
# Initializing
login_url = "https://orbis4.bvdinfo.com/"
form_data = {"user": "WBG_IFC", "pw": "Global Markets"}

browser = webdriver.Chrome()

login_orbis(browser)

for num in range(5,12):
    select_file(browser, 'EMPEA_raw_data', num)
    
    total_page_num = (int(browser.find_element_by_css_selector('body > div.viewport.main > div.website > div.content > div > div.title > h2 > span').text[:4])-1)//100+1
    
    select_score(browser,total_page_num,1)
    company_mapping = create_mapping(browser,total_page_num)
    
    # Go to Results Page
    browser.find_element_by_css_selector('body > div.viewport.main > div.website > div.pre-content > ul > li:nth-child(1) > a').click()
    time.sleep(2)
    company_data = data_scraping(browser)
    
    final_data = company_mapping.merge(company_data,left_on='mapped_bvdId',right_on='BvD ID number ', how='left')
    final_data.to_csv('Mapped_data.csv', mode='a', index=False)

=======
# -*- coding: utf-8 -*-
"""
Created on Mon Jun 25 08:57:10 2018

@author: sqian
"""


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup as soup
from bs4 import SoupStrainer
from lxml import html
import numpy as np
import time
import pandas as pd
import win32com.client as win32
import os


def login_orbis(browser):
    # Turn on the broswer
    browser.get(login_url)
    # Login Page
    username = browser.find_element_by_name("user")
    password = browser.find_element_by_name("pw")
    username.send_keys(form_data["user"])
    password.send_keys(form_data["pw"])
    login_button = browser.find_element_by_class_name("ok")
    login_button.click()
    try:
        restart_button = browser.find_element_by_xpath("//input[@class='button ok']")
        restart_button.click()
    except:
        pass
                                                    

def hard_refresh(browser,year,start_page):
    browser.close()
    
    browser = webdriver.Chrome()
    browser.get(login_url)
    username = browser.find_element_by_name("user")
    password = browser.find_element_by_name("pw")
    username.send_keys(form_data["user"])
    password.send_keys(form_data["pw"])
    while 1:
        try:
            login_button = browser.find_element_by_class_name("ok")
            login_button.click()
            break
        except:
            browser.close()
            browser = webdriver.Chrome()
            browser.get(login_url)
    
    login_orbis(browser,year)
    while 1:
        try:
            page_input = browser.find_elements_by_css_selector("ul.navigation > li > input")[0]
            page_input.clear()
            page_input.send_keys(str(start_page))
            page_input.send_keys(Keys.RETURN)
            break
        except:
            continue
    refresh_stuck = visible_in_time(browser, '#resultsTable > tbody > tr > td.scroll-data > div > table > tbody > tr:nth-child(1) > td:nth-child(1)', 20)
    if refresh_stuck == False:
        return hard_refresh(browser,year,start_page)
    else:
        return browser
 
    
def visible_in_time(browser, address, time):
    try:
        WebDriverWait(browser, time).until(EC.presence_of_element_located((By.CSS_SELECTOR, address)))
        return True
    except TimeoutException:
        return False


def split_table(file_name):
    datatable = pd.read_csv(file_name)
    for i in range(len(datatable)//1000+1):
        datatable.loc[i*1000:(i+1)*1000].to_csv(file_name[:-4] + '_' + str(i) + '.csv', mode='w', index=True, sep='\t')


def select_score(browser, total_page_num, start_page):
    # Go to Page 1
    page_input = browser.find_element_by_xpath('//li/input[@type="number"]')
    page_input.clear()
    page_input.send_keys(start_page)
    page_input.send_keys(Keys.RETURN)
    print("Start Mapping!")
    time.sleep(1)
    for page in range(start_page-1,total_page_num):
        innerHTML = browser.execute_script("return document.body.innerHTML")
        tree = html.fromstring(innerHTML)
        orbis_result = tree.xpath('//*[@id="matchedSelected"]/@data-selected')
        orbis_content = tree.xpath('//*[@id="matchedSelected"]/@data-matched')
        for id in range(len(orbis_result)):
            if orbis_result[id] == '[]' and orbis_content[id] != '[]':
                nation = browser.find_element_by_xpath('//tbody/tr[@data-id={0}]/td[@data-id="Country"]/div'.format(id+page*100)).text
                item_expand = browser.find_element_by_xpath('//tbody/tr[@data-id={0}]/td/label'.format(id+page*100))
                item_expand.click()
                try:
                    score = browser.find_element_by_xpath('//*[@id="Template" and not(string(@style))]/td[7]/div').text
                    nation_popped = browser.find_element_by_xpath('//*[@id="Template" and not(string(@style))]/td[4]/div').text
                    if nation == nation_popped or len(nation) == 3:
                        browser.find_element_by_xpath('//*[@id="Template" and not(string(@style))]/td[1]/label').click()
                        print(id+page*100)
                        time.sleep(1.5)
                    else:
                        print("Mapping result rejected!")
                except:
                    continue
        browser.find_element_by_xpath('/html/body/div[2]/div[1]/div[4]/div/form/div[2]/ul/li[12]/img').click()


def create_mapping(browser, total_page_num):
    # Generate Mapping Table
    page_input = browser.find_element_by_xpath('//li/input[@type="number"]')
    page_input.clear()
    page_input.send_keys(1)
    page_input.send_keys(Keys.RETURN)
    time.sleep(1)
    print("Start Generating Mapping Table!")
    
    company_mapping = pd.DataFrame(columns=['original_name','country_code','mapped_name','mapped_bvdId','mapping_score'])
    
    for page in range(0,total_page_num):
        innerHTML = browser.execute_script("return document.body.innerHTML")
        tree = html.fromstring(innerHTML)
        data_temp = pd.DataFrame()
        data_temp['original_name'] = tree.xpath('//td[@data-id="Name"]/div/text()')
        data_temp['country_code'] = [x.text for x in tree.xpath('//td[@data-id="Country"]/div')]
        data_temp['mapped_name'] = [x.strip() for x in tree.xpath('//td[@id="matchedSelected"]/div/text()')]
        result_list = [eval(x) for x in tree.xpath('//td[@id="matchedSelected"]/@data-selected')]
        id_list = []
        for x in result_list:
            try:
                id_list.append(x[0]['BvDId'])
            except:
                id_list.append("")
        data_temp['mapped_bvdId'] = id_list
        data_temp['mapping_score'] = tree.xpath('//td[@id="matchedScore"]/div/@class')
        company_mapping = pd.concat([company_mapping,data_temp])
        browser.find_element_by_xpath('/html/body/div[2]/div[1]/div[4]/div/form/div[2]/ul/li[12]/img').click()
    
    # Output mapping table
    company_mapping.to_csv('mapping.csv', mode='a', index=False)
    return company_mapping


def data_scraping(browser):
    # Download Detailed Data
    select_view = browser.find_element_by_css_selector('div.menuViewContainer > div.menuView > ul > li > a')
    select_view.click()  
    if visible_in_time(browser, 'span.name.clickable[title="EMPEA"]', 30):
        view_year = browser.find_element_by_css_selector('span.name.clickable[title="EMPEA"]')
        view_year.click()
    
    innerHTML = browser.execute_script("return document.body.innerHTML")
    page_soup = soup(innerHTML, "html.parser")
    columns = page_soup.find('table', {'class': 'scroll-header'}).find_next('tr')
    label_info = columns.find_all('div', class_='column-label')
    column_names = []
    for x in label_info:
        try:
            column_names.append(x.span['data-fulllabel'] + ' <' + x.find_all('span')[1]['data-full-configuration'] + '>')
        except:
            column_names.append(x.span['data-fulllabel'])
    company_data = pd.DataFrame()
    company_names = []
    for x in column_names:
        company_data[x] = []
    
    per_page = 100
    total_companies = int(page_soup.find('td',{'class':'grand-total'}).text.replace(',',''))
    total_pages =  total_companies // per_page + 1 # Number of pages of data to retrieve
    page_done = 0
    
# =============================================================================
#     page_input = browser.find_elements_by_css_selector("ul.navigation > li > input")[0]
#     page_input.clear()
#     page_input.send_keys(str(1))
#     page_input.send_keys(Keys.RETURN)
# =============================================================================
    if visible_in_time(browser,'#resultsTable > tbody > tr > td.scroll-data > div > table > tbody > tr:nth-child(1) > td:nth-child(1)',20):
        pass
    else:
        print('Timeout!')
        exit()
    time.sleep(4)
    
    while page_done < total_pages:
        innerHTML = browser.execute_script("return document.body.innerHTML")
        tree = html.fromstring(innerHTML)
        page_num = int(tree.xpath('//ul[@class="navigation"]/*/span[@class="currentPage" and text() != "..." ]/text()')[0])
        if page_num == page_done+1 and visible_in_time(browser,'#resultsTable > tbody > tr > td.scroll-data > div > table > tbody > tr:nth-child(1) > td:nth-child(1)',20):
            print("Page {0} retrieved!".format(page_num))
            tree = tree.cssselect('#resultsTable')[0]
            if company_names == []:
                company_names = [x.text for x in tree.xpath('//span[@class="ellipsis"]/a[@href="#"]')]
            else:
                company_names += [x.text for x in tree.xpath('//span[@class="ellipsis"]/a[@href="#"]')]
            data_points = tree.xpath('//td[@class="scroll-data"]/div/table/tbody/tr/descendant::*/text()')
            if page_num == total_pages:
                num_on_page = len(tree.xpath('//span[@class="ellipsis"]/a[@href="#"]'))
            else:
                num_on_page = per_page
            data = np.array_split(data_points, num_on_page)                
            company_data = pd.concat([company_data,pd.DataFrame(data,columns=column_names)])
            
            page_done += 1
            print("Page {0} finished!".format(page_num))
            if page_num != total_pages:
                browser.find_element_by_xpath("//img[@data-action='next']").click()
    
    company_data.insert(0,"company_name",company_names)
    return company_data


def select_file(browser, file_name, file_id):
    # Turn to the page of Tools
    batch_search_page = "https://orbis4.bvdinfo.com/version-2018621/orbis/1/Companies/BatchSearch/Start"
    browser.get(batch_search_page)
    time.sleep(2)
    browser.find_element_by_id('upload-now').click()
    browser.find_element_by_css_selector('body > div.viewport.main > div.website > div.content > div > div.batchWidget > div > div > form > div.view > div:nth-child(1) > input.hidden').send_keys(os.getcwd()+'/' + file_name + '_{0}.csv'.format(file_id))
    browser.find_element_by_css_selector('dl.mapping-options > dd:nth-child(3) > label').click()
    browser.find_element_by_css_selector('div.batchWidget > div > div > form > div.buttons > div > a.button.ok').click()
    if visible_in_time(browser,'#CountDown',20):                  
        while 1:
            search_process = browser.find_element_by_css_selector('#CountDown').text.split('/')
            if search_process[0] == search_process[1]:
                time.sleep(5)
                break
            else:
                time.sleep(1)
                                                              

###### Main Function ######
# Initializing
login_url = "https://orbis4.bvdinfo.com/"
form_data = {"user": "WBG_IFC", "pw": "Global Markets"}

browser = webdriver.Chrome()

login_orbis(browser)

for num in range(5,12):
    select_file(browser, 'EMPEA_raw_data', num)
    
    total_page_num = (int(browser.find_element_by_css_selector('body > div.viewport.main > div.website > div.content > div > div.title > h2 > span').text[:4])-1)//100+1
    
    select_score(browser,total_page_num,1)
    company_mapping = create_mapping(browser,total_page_num)
    
    # Go to Results Page
    browser.find_element_by_css_selector('body > div.viewport.main > div.website > div.pre-content > ul > li:nth-child(1) > a').click()
    time.sleep(2)
    company_data = data_scraping(browser)
    
    final_data = company_mapping.merge(company_data,left_on='mapped_bvdId',right_on='BvD ID number ', how='left')
    final_data.to_csv('Mapped_data.csv', mode='a', index=False)

print("Done!")
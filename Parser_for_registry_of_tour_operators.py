#!/usr/bin/env python3
# -*- coding: utf8 -*-

import sys
import urllib.request
import pandas as pd
import numpy as np
import csv
from pandas import ExcelWriter
import re
import openpyxl
from collections import namedtuple
from bs4 import BeautifulSoup
from traceback import print_exc
import lxml.etree as etree


import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.firefox.options import Options


Operator = namedtuple('Operator', ['item_title', 'name', 'href', 'url', 'INN', 'OGRN', 'authenticity', 'status'])

toperators = []

def init_driver():
    print("Wait 10 seconds...")
    options = Options()
    options.add_argument("--headless")    
    driver = webdriver.Firefox(firefox_options=options, executable_path='C:\Geckodriver\geckodriver.exe')
    driver.wait = WebDriverWait(driver, 5)    
    return driver


def get_list(driver):
    try:
        try:
            driver.get("https://www.russiatourism.ru/operators/")
        except TimeoutException:
            print("Can't find the website")
            exit()
            
        # Click "add filtr"
        try:
            button = driver.wait.until(EC.element_to_be_clickable(
                (By.CLASS_NAME, "button-add-filter")))
            button.click()
        except TimeoutException:
            print("Can't find filter button")
            exit()
            
        # Click "choose from the list"
        try:
            button = driver.wait.until(EC.element_to_be_clickable(
                (By.ID, "sg_type-selectized"))) 
            button.click()
        except TimeoutException:
            print("Can't find button opening the list of filters")   
            exit()
            
        # Click "turoperator's activity"             
        try:   
            button = driver.wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//div[@data-value='type-turism']"))) 
            button.click()           
        except TimeoutException:
            print("Can't find 'Сфера туроператорской деятельности'")
            exit()
            
        # Click "choose turoperator's activity"
        try:
            button = driver.wait.until(EC.element_to_be_clickable(
                (By.ID, "type-turism-selectized")))  
            button.click() 
        except TimeoutException:
            print("Can't find button opening the list of turism types")
            exit()
            
        # Click "international outer tourism"
        try:
            button = driver.wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//div[@data-value='out']"))) 
            button.click()     
        except TimeoutException:
            print("Can't find 'Международный выездной туризм'") 
            exit()
                   
        # Click "apply"
        try:
            button = driver.wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//input[@class='btn btn-submit']"))) 
            button.click()  
        except TimeoutException:
            print("Can't find apply-button")
            exit()
        
        time.sleep(5)   
        
        # Getting page content
        soup = BeautifulSoup(driver.page_source, 'lxml')     
        # Getting amount of operators
        operators_count_text = soup.find('div', {'class' : 'search-result_title'}).text
        operators_count = int(operators_count_text.split(":")[-1])
        
        # First page
        make_list(soup)
        # Each page has "number" operators
        number = count_operators_on_page(soup)
        i = number
        
        '''While there is the button "next page" do parsing'''
        while EC.element_to_be_clickable((By.XPATH, "//[@class='last-page pull-right']")):            
            button = driver.wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//a[@class='last-page pull-right']"))) 
            button.click()
            time.sleep(1.5)
            i += number
            if (i / operators_count * 100) <= 100:
                print(f'Parsing = {round(i / operators_count * 100)}%')
            else:
                print(f'Parsing = {100}%')
            
            soup = BeautifulSoup(driver.page_source, 'lxml')        
            make_list(soup)    
            #time.sleep(1)                
    except TimeoutException:
        print("The first part is complete")  
    
      
def get_html(url):
    response = urllib.request.urlopen(url)
    return response.read() 

def has_cyrillic(text):
    search = re.search(r'[а-яА-ЯёЁ]', text)
    return bool(search)
        
def get_additional_info(toperators):
    '''Getting operators URL, INN, OGRN'''
    i = 0
    percentage = 0
    print(len(toperators))
    while i < len(toperators):
        try:
            soup = BeautifulSoup(get_html(toperators[i].href), 'lxml') 
            
            divs_set = soup.find_all('div', {'class' : 'b_inner__regis_item'})
                    
            p = 0
            if len(divs_set) != 0:            
                while p < len(divs_set):            
                    if divs_set[p].text != 'Адрес официального сайта в сети "Интернет":':
                        p += 1
                    else:
                        url = divs_set[p+1].text
                        if (has_cyrillic(url) == True):
                            if (('нет' in url) or ('не ' in url)):
                                got_url = 'None'                                
                            else:
                                got_url = url
                        else:
                            if ('--' in url):
                                got_url = 'None'
                            else:                              
                                got_url = url 
                        break                                   
            else: 
                print(f"Can't find additional info on page {toperators[i].href}")
                got_url = "404" 
                
            p = 0
            if len(divs_set) != 0:            
                while p < len(divs_set):            
                    if divs_set[p].text != 'ИНН:':
                        p += 1
                    else:
                        got_inn = divs_set[p+1].text 
                        break                                   
            else: 
                print(f"Can't find additional info on page {toperators[i].href}")
                got_inn = "404"
                
            p = 0
            if len(divs_set) != 0:            
                while p < len(divs_set):            
                    if divs_set[p].text != 'ОГРН:':
                        p += 1
                    else:
                        got_ogrn = divs_set[p+1].text 
                        break                                   
            else: 
                print(f"Can't find additional info on page {toperators[i].href}")
                got_ogrn = "404"
                
            toperators[i] = toperators[i]._replace(url = got_url, INN = got_inn, OGRN = got_ogrn)
            if percentage < round((i / len(toperators)) * 100):
                percentage = round((i / len(toperators)) * 100)
                print(f'Additional info parsing = {percentage}%')
            else:
                pass
            i += 1  
        except:
            print(f"Page with additional info for '{toperators[i].name}' not found")
            got_url = "404" 
            got_inn = "404"
            got_ogrn = "404"
            toperators[i] = toperators[i]._replace(url = got_url, INN = got_inn, OGRN = got_ogrn) 
            if percentage < round((i / len(toperators)) * 100):
                percentage = round((i / len(toperators)) * 100)
                print(f'Additional info parsing = {percentage}%')
            else:
                pass            
            i += 1  
    
   
def make_list(soup):
    '''Replenishment of the list'''
    table = soup.find('div', {'class' : 'col-md-12 search-result'})
            
    for row in table.find_all('div', {'class' : 'search-result_item'}):
        string = row.find_all('div', {'class' : 'search-result_item_title'})        
               
        toperators.append(Operator(row.find('div', class_ = "search-result_item_title").text, 
                                   row.find('a', class_ = "search-result_item_link").text, 
                                   "https://www.russiatourism.ru" + row.find('a', class_ = "search-result_item_link").get('href'),
                                   None,
                                   None,
                                   None,
                                   None, 
                                   None)) 
   
    
def form_list(toperators):
    '''Converts URL's to a common type, fix domens mistakes (,com -> .com)'''
    df = pd.DataFrame(toperators)
    
    domains_df = pd.read_csv('domains1.csv')
    
    for i in range(len(df)):        
        if (df.loc[i, 'name'][0] == ' '):
            df.loc[i, 'name'] = df.loc[i, 'name'][1:]
            
        if (('\r' in df.loc[i, 'name']) or ('\n' in df.loc[i, 'name']) or ('\t' in df.loc[i, 'name'])):
            df.loc[i, 'name'] = df.loc[i, 'name'].replace('\r', ' ')
            df.loc[i, 'name'] = df.loc[i, 'name'].replace('\n', ' ')
            df.loc[i, 'name'] = df.loc[i, 'name'].replace('\t', ' ')
            
        df.loc[i, 'name'] = df.loc[i, 'name'].replace('"', "'")
        df.loc[i, 'name'] = df.loc[i, 'name'].replace('«', "'")
        df.loc[i, 'name'] = df.loc[i, 'name'].replace('»', "'")
        
        if ',' in df.loc[i, 'url']:
            for j in range(len(domains_df)):
                if f",{domains_df.loc[j, 'name'].split('.')[1]}" in df.loc[i, 'url']:
                    df.loc[i, 'url'] = df.loc[i, 'url'].replace(',', '.')
            print(df.loc[i, 'url'])
        # Replace russian [эс] among english characters with english [си]    
        if (('с' in df['url'][i]) and (has_cyrillic(df['url'][i].split('с')[0]) == False)):
            df['url'][i] = df['url'][i].replace('с', 'c')
            
        if (df['url'][i][0:7] == 'http://'):
            df['url'][i] = df['url'][i].replace('http://','')
            
        if (df['url'][i][0:4] != 'www.'):            
            if ('None' not in df['url'][i]) and ('404' not in df['url'][i]):
                df['url'][i] = 'www.' + df['url'][i]
                    
    df.to_csv('testWING1.csv', encoding='utf-8', index=False)
    return df
    
    
def saveCSV(df, OutputFile):
    df.to_csv(f'{OutputFile}.csv', encoding='utf-8', index=False)
 
    
def saveEXCEL(df, OutputFile):    
    writer = pd.ExcelWriter(f'{OutputFile}.xlsx')    
    df.to_excel(writer,'Sheet1')    
    writer.save()     

    
def saveXML(df, OutputFile):
    root = etree.Element('Operators_list')
    
    for row in df.iterrows():    
        Operator = etree.SubElement(root, 'Operator') 
        item_title = etree.SubElement(Operator, 'item_title')    
        name = etree.SubElement(Operator, 'name')  
        url = etree.SubElement(Operator, 'url')
        additional_url = etree.SubElement(Operator, 'additional_url')
        INN = etree.SubElement(Operator, 'INN')
        OGRN = etree.SubElement(Operator, 'OGRN')
        authenticity = etree.SubElement(Operator, 'authenticity')        
        status = etree.SubElement(Operator, 'status')
        
        item_title.text = str(row[1]['item_title'])
        name.text = str(row[1]['name'])
        url.text = str(row[1]['url'])
        additional_url.text = str(row[1]['additional_url'])
        INN.text = str(row[1]['INN'])
        OGRN.text = str(row[1]['OGRN'])
        authenticity.text = str(row[1]['authenticity'])
        status.text = str(row[1]['status']) 
        
        handle = etree.tostring(root, pretty_print=True, encoding='utf8', xml_declaration=True).decode('utf8')
                
    with open(f'{OutputFile}.xml', 'w', newline=None, encoding='utf8') as f:
        f.write(handle) 
        f.close()
        
        
def saveXML_attribute_style(df, OutputFile):
    root = etree.Element('Operators_list')
    
    for row in df.iterrows():
        Operator = etree.Element("Operator")
        Operator.set('item_title', str(row[1]['item_title']))
        Operator.set('name', str(row[1]['name']))
        Operator.set('url', str(row[1]['url']))
        Operator.set('additional_url', str(row[1]['additional_url']))
        Operator.set('INN', str(row[1]['INN']))
        Operator.set('OGRN', str(row[1]['OGRN']))
        Operator.set('authenticity', str(row[1]['authenticity']))
        Operator.set('status', str(row[1]['status']))
        root.append(Operator)
        handle = etree.tostring(root, pretty_print=True, encoding='utf8', xml_declaration=True).decode('utf8')
                
    with open(f'{OutputFile}.xml', 'w', newline=None, encoding='utf8') as f:
        f.write(handle) 

def split_urls(row):
    if ', ' in row[1]['url']:
        return row[1]['url'].split(', ')
    else:
        if ' ' in row[1]['url']:
            return row[1]['url'].split(' ')
        else:
            return row[1]['url']


def saveXML(df, OutputFile):
    root = etree.Element('Operators_list')
    
    for row in df.iterrows():    
        Operator = etree.SubElement(root, 'Operator')    
        name = etree.SubElement(Operator, 'name')
        if (type(split_urls(row)) != type([])):
            url = etree.SubElement(Operator, 'url')
            url.text = str(split_urls(row))
        else:
            for i in range(len(split_urls(row))):
                if (split_urls(row)[i] != ''):
                    url = etree.SubElement(Operator, 'url')
                    url.text = str(split_urls(row)[i])
                else:
                    pass
       
        name.text = str(row[1]['name'])
        name.set('item_title', str(row[1]['item_title']))
        name.set('INN', str(row[1]['INN']))
        name.set('OGRN', str(row[1]['OGRN']))
        if (str(row[1]['authenticity']) != 'nan'):
            name.set('authenticity', str(row[1]['authenticity']))
        if (str(row[1]['status']) != 'nan'):
            name.set('status', str(row[1]['status']))
        
        handle = etree.tostring(root, pretty_print=True, encoding='utf8', xml_declaration=True).decode('utf8')
                
    with open(f'{OutputFile}.xml', 'w', newline=None, encoding='utf8') as f:
        f.write(handle) 
        f.close()        
      
        
def count_operators_on_page(soup): 
    '''Count the number of operators on one page'''
    count = len(soup.find_all('div', {'class' : 'search-result_item'}))
    return count

        
def main():
    print('How to save: press "0" for CSV, press "1" for EXCEL, press "2" for BOTH')
    saver = input()
    if (saver != '0') & (saver != '1') & (saver != '2'):
        print('Please, be more careful next time and press "0" for CSV, press "1" for EXCEL or press "2" for BOTH')
        exit()        
    print('Enter the name of the file without filename extension (e.g. MyFile)')
    OutputFile = input()
    
    ## Parsing
    driver = init_driver()
    get_list(driver) 
    driver.quit()
    get_url(toperators)
    
    if (saver == '0') or (saver == '2'):
        ## Save CSV
        saveCSV(df, OutputFile)
        if (saver == '2'):
            ## Save EXCEL    
            saveEXCEL(df, OutputFile)
    else:
        ## Save EXCEL        
        saveEXCEL(df, OutputFile)        
        
    
    print('Do you want to count approximate number of operators without URL? Press "Y"/"N" for Yes/No')
    answer = input()
    
    if (answer == 'Y') or (answer == 'y'):
        k = 0
        count_of_none_url_operators = 0
        while k < len(toperators):
            textik = str(toperators[k].url)
            if (textik.find('.') != -1) or (textik.find('www.') != -1) or (textik.find('.ru') != -1) or (textik.find('.com') != -1) or (textik.find('.org') != -1):
                k += 1            
            else:
                count_of_none_url_operators += 1
                k += 1  
        print(f'Number of operators without URL = {count_of_none_url_operators}') 

if __name__ == '__main__':
    main()
    

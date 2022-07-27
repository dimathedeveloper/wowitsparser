import requests
from bs4 import BeautifulSoup
import pandas as pd
import time

def load_page(url =  'http://www.list-org.com/company/5553931'):
    headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) \
        AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.120 Safari/537.36'}

    if not url.startswith('http'):
        return None

    response = requests.get(url, headers=headers)


    if not response.ok:
        return None

    return response.text

##### Search page {
def load_search_page(url):

    if not url.startswith('http'):
        return None
    headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) \
        AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.120 Safari/537.36'}
    response = requests.get(url, headers=headers)

    if not response.ok:
        return None

    return response.text

def parse_search_page(text):
    soup = BeautifulSoup(text, 'html.parser')
    page = soup.find('div', {'class' : 'content'}).text
    
    if ('Найдено 0 организаций' in page):
        return None

    orgs = soup.find('div', {'class' : 'org_list'})
    urls = orgs.find_all('label')

    # Sometimes we have more then one org's in org_list
    urlsDict = []
    for item in urls:
        urlsDict.append(item.find('a').get('href'))
        
    return urlsDict
##### Search page }


def get_company_data(text):
    soup = BeautifulSoup(text, 'html.parser')
    company_data = soup.find_all('div', {'class': 'main'})

    return company_data


def get_need_data(text):
    company_info_dict = dict.fromkeys(['State', 'People_amount', 'Core', 'Income'])

    general_text = text[0]

    table = general_text.find('table', {'class' : 'table table-sm'})
    table_items = table.find_all('tr')

    cState = ""
    cPAmount = ""
    cCore = ""
    cIncome = ""

    
    for item in table_items:
        splited_item = item.text.split(':', 1)
        # State
        if (splited_item[0] == 'Статус'):
            cState = splited_item[1]
        # cPAmount
        if (splited_item[0] == 'Численность персонала'):
            cPAmount = str(int(splited_item[1]))
            
    allCards = general_text.find_all('div', {'class' : 'card w-100 p-1 p-lg-3 mt-2'})
    for card in allCards:
        try:
            cardTitle = card.find('h6', {'class' : 'd-flex card-title'})
            titleName = cardTitle.find('div').text
        except:
            continue
        if (titleName == "Виды деятельности:"): # Find OKVED
            for item in card:
                sItem = item.text.split(':', 1)
                if (sItem[0] == 'Основной (по коду ОКВЭД ред.2)'):
                    cCore = sItem[1].split('\n', 1)[0]
            pass
        if ("Сведения о доходах и расходах" in titleName): # Find Income
            tableIncome = card.find('table')
            if ("Доходы" in tableIncome.text):
                values = tableIncome.find('td', {'class' : 'nwra'})
                for item in values:
                    cIncome = item.text
                    break


    ### End
    company_info_dict['State'] = cState
    company_info_dict['People_amount'] = cPAmount
    company_info_dict['Core'] = cCore
    company_info_dict['Income'] = cIncome

    return company_info_dict
    
#######################################################



def getOrgData(searchUrl, mainUrl):
    haveCaptha = True

    while (haveCaptha):
        try:
            sText = load_search_page(searchUrl)
            
            needUrl = parse_search_page(sText)

            if (needUrl == None):
                return None

            company_info_arr = []
            for url in needUrl:
                text = load_page(mainUrl + url)
                company_data = get_company_data(text)
                company_info_dict = get_need_data(company_data)
                company_info_arr.append(company_info_dict)

            haveCaptha = False
        except:
            print("Check capcha!")
            time.sleep(10)

    
    # Здесь возвращаются нужные для экселя данные(Статус, Численность, Основной и Доходы соответственно)
    #return company_info_dict
    return company_info_arr


MAIN_URL = 'http://www.list-org.com'
SEARCH_URL = 'https://www.list-org.com/search?type=inn&val='

import math

def main():
    # Get .xlsx
    data = pd.read_excel('./dlya_TZ_parsing2.xlsx')
    pd.options.display.max_colwidth = 500
    # Get INN's
    INNarr = data['ИНН / ЕДРПОУ(Укр) пользователя'].to_list()

    # Main cycle {
    for i in range (len(INNarr)):
        # Not null check
        if (math.isnan(INNarr[i])):
            continue

        text = getOrgData(SEARCH_URL + str(int(INNarr[i])), MAIN_URL)
        # Если страница не загружается...
        if (text) is None:
            # data.loc[i,'Статус'] = 'ERROR'
            # data.loc[i,'Численность персонала'] = 'ERROR'
            # data.loc[i,'Основной (по коду ОКВЭД ред.2)'] = 'ERROR'
            # data.loc[i,'Доходы 2020 год'] = 'ERROR'
            #Цикл продолжается
            continue
        # ................................
        # Execute state
        print(INNarr[i], f'Complete {i + 1}/{len(INNarr)}')

        state = ""
        amount = ""
        core = ""
        income = ""
        
        for info in text:
            if len(info['State']) > 0:
                state += (info['State'] + ',')
            if len(info['People_amount']) > 0:
                amount += (info['People_amount'] + ',')
            if len(info['Core']) > 0:
                core += (info['Core'] + ',')    
            if len(info['Income']) > 0:
                income += (info['Income'] + ',')


        data.loc[i,'Статус'] = state
        data.loc[i,'Численность персонала'] = amount
        data.loc[i,'Основной (по коду ОКВЭД ред.2)'] = core
        data.loc[i,'Доходы 2020 год'] = income
    # Main cycle }

    # Save frame on directory
    data.to_excel('./gotovo.xlsx')
        


if __name__ == '__main__':
    main()

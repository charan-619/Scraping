from playwright.sync_api import sync_playwright
import time
from bs4 import BeautifulSoup
import re
import pandas as pd


def get_financial_results(page,company):
        page.locator("a#afi[data-toggle='collapse'][data-parent='#accordion']").click()
        time.sleep(3)
        page.locator("a#l61").click()
        time.sleep(6)
        soup = BeautifulSoup(page.content(), 'html.parser')
        element = soup.find_all('table', class_ ="ng-binding")[0]
        row_list= list()
        tr_heads = element.find('thead').find_all('tr')
        tr_body = element.find('tbody').find_all('tr')
        columns = list()
        for tr_tag in tr_body:
            l = list()
            try:
                td_tags = tr_tag.find_all('td', class_='tdcolumn')
                for td_tag in td_tags:
                    l.append(td_tag.text.strip())
            except:
                continue
            row_list.append(l)
        for th_tag in tr_heads:
            td_tags = th_tag.find_all('td', class_='tableheading')
            for td_tag in td_tags:
                columns.append(td_tag.text.strip())
        dataframe1 = pd.DataFrame(row_list,columns=columns)
        dataframe1 = dataframe1[:-3]
        dataframe1 =dataframe1.drop(dataframe1.index[0])
        page.get_by_role("link", name="Annual Trends").click()
        time.sleep(3)
        soup1 = BeautifulSoup(page.content(), 'html.parser')
        element1 = soup1.find_all('table', class_ ="ng-binding")[3]
        row_list1= list()
        tr_heads1 = element1.find('thead').find_all('tr')
        tr_body1 = element1.find('tbody').find_all('tr')
        columns1 = list()
        for tr_tag1 in tr_body1:
            l1 = list()
            try:
                td_tags1 = tr_tag1.find_all('td', class_='tdcolumn')
                for td_tag1 in td_tags1:
                    l1.append(td_tag1.text.strip())
            except:
                continue
            row_list1.append(l1)
        for th_tag1 in tr_heads1:
            td_tags1 = th_tag1.find_all('td', class_='tableheading')
            for td_tag1 in td_tags1:
                columns1.append(td_tag1.text.strip())
        dataframe2 = pd.DataFrame(row_list1,columns=columns1)
        dataframe2 = dataframe2[:-3]
        dataframe2 =dataframe2.drop(dataframe2.index[0])
        with pd.ExcelWriter(f'{company}_financial_results1.xlsx', engine='xlsxwriter') as writer:
            dataframe1.to_excel(writer, sheet_name='Quaterly trends', index=False)
            dataframe2.to_excel(writer, sheet_name='Annual trends', index=False)


def main():
    with sync_playwright() as p:
        browser = p.firefox.launch(headless=False)
        page = browser.new_page()
        link = 'https://www.bseindia.com/stock-share-price/reliance-industries-ltd/reliance/500325/'
        page.goto(link, timeout=60000)
        time.sleep(2)
        l = link.split('/')
        company = l[4].strip().capitalize()
        get_financial_results(page, company)
        


if __name__ == "__main__":
    main()
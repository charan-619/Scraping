from playwright.sync_api import sync_playwright
import time
from bs4 import BeautifulSoup
import re
import pandas as pd


def get_equity_data(page, company):
    time.sleep(3)
    dict1 = dict()
    soup = BeautifulSoup(page.content(), 'html.parser')

    div_elements = soup.find_all('div', class_='col-lg-13')
    for element in div_elements:
        try:
            tr_tags = element.find('tbody').find_all('tr')
        except:
            continue
        for tr_tag in tr_tags:
            try:
                key = tr_tag.find('td', class_='textsr').text.strip()
                value = tr_tag.find('td', class_='textvalue ng-binding').text.strip()
                dict1[key] = value
            except:
                continue
    
    row_list = []
    data = dict1
    row_list.append(data.values())
    df = pd.DataFrame(row_list, columns=data.keys())
    return df

def financial_results(page,company):
    page.locator("a#afi[data-toggle='collapse'][data-parent='#accordion']").click()
    time.sleep(2)
    page.locator("a#l61").click()
    time.sleep(3)
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
    return dataframe1, dataframe2


def get_board_meetings(page,company):
    page.locator('a#ame').click()
    time.sleep(3)
    page.locator('a#l71').click()
    time.sleep(6)
    soup = BeautifulSoup(page.content(), 'html.parser')
    element = soup.find_all('table', class_ = "ng-scope")[1]
    row_list= list()
    tr_heads = element.find('thead').find_all('tr')
    tr_body = element.find('tbody').find_all('tr')
    columns = list()
    for tr_tag in tr_body:
        l = list()
        try:
            td_tags = tr_tag.find_all('td', class_='tdcolumn ng-binding')
            for td_tag in td_tags:
                l.append(td_tag.text.strip())
        except:
            continue
        row_list.append(l)
    for th_tag in tr_heads:
        td_tags = th_tag.find_all('td', class_='tableheading')
        for td_tag in td_tags:
            columns.append(td_tag.text.strip())
    dataframe = pd.DataFrame(row_list,columns=columns)
    return dataframe



def get_corpgov_data( page, company):
    page.get_by_role("link", name="Corporate Governance").click()
    time.sleep(3)
    soup = BeautifulSoup(page.content(), 'html.parser')
    element = soup.find_all('table', class_='ng-scope')[4]
    row_list= list()
    tr_body = element.find('tbody').find_all('tr')
    # tr_heads = element.find('thead').find_all('tr')

    for tr_tag in tr_body:
        l = list()
        try:
            td_tags = tr_tag.find_all('td', class_='ng-binding')
            for td_tag in td_tags:
                l.append(td_tag.text.strip())
            
        except:
            continue
        row_list.append(l)
    needed_row_list = list()
    for lis in row_list:
        row_list1 = [lis[0],lis[1],lis[2],lis[4],lis[9],lis[12],lis[13]]
        needed_row_list.append(row_list1)
    column_data1 = ['Sr','Title (Mr/Ms)','Name of the Director','Category','Current status','Initial Date of Appointment','Date of Re-appointment']
    dataframe = pd.DataFrame(needed_row_list,columns=column_data1)
    return dataframe

def get_shareholder_meetings(page, company):
    # page.locator('a#ame').click()
    # time.sleep(3)
    page.locator('a#l72').click()
    time.sleep(6)
    soup = BeautifulSoup(page.content(), 'html.parser')
    element = soup.find_all('table', class_ = "ng-scope")[1]
    row_list= list()
    tr_heads = element.find('thead').find_all('tr')
    tr_body = element.find('tbody').find_all('tr')
    columns = list()
    for tr_tag in tr_body:
        l = list()
        try:
            td_tags = tr_tag.find_all('td', class_='tdcolumn ng-binding')
            for td_tag in td_tags:
                l.append(td_tag.text.strip())
        except:
            continue
        row_list.append(l)
    for th_tag in tr_heads:
        td_tags = th_tag.find_all('td', class_='tableheading')
        for td_tag in td_tags:
            columns.append(td_tag.text.strip())
    dataframe = pd.DataFrame(row_list,columns=columns)
    return dataframe



def get_peer_group(page,company):
    page.get_by_role("heading", name="Peer Group").get_by_role("link").click()
    time.sleep(3)
    soup = BeautifulSoup(page.content(), 'html.parser')
    element1 = soup.find('div',id = "qtly")
    element2 = element1.find_all('table')[0]
    element = element2.find('table')
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
    page.get_by_role('link', name="Annual Trends").click()
    time.sleep(3)
    soup1 = BeautifulSoup(page.content(), 'html.parser')
    element3 = soup1.find('div',id = "ann")
    element4 = element3.find_all('table')[0]
    element5 = element4.find('table')
    row_list1= list()
    tr_heads1 = element5.find('thead').find_all('tr')
    tr_body1 = element5.find('tbody').find_all('tr')
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
    
    page.get_by_role('link', name="Bonus & Dividends").click()
    time.sleep(3)
    soup1 = BeautifulSoup(page.content(), 'html.parser')
    element6 = soup1.find('div',id = "bnd")
    element7 = element6.find_all('table')[0]
    element8 = element7.find('table')
    row_list2= list()
    tr_body2 = element8.find('tbody').find_all('tr')
    for tr_tag2 in tr_body2:
        l2 = list()
        try:
            td_tags2 = tr_tag2.find_all('td', class_='tdcolumn')
            for td_tag2 in td_tags2:
                l2.append(td_tag2.text.strip())
        except:
            continue
        row_list2.append(l2)
    row_list3 = row_list2[7:]
    dataframe3 = pd.DataFrame(row_list3)

    return dataframe1,dataframe2,dataframe3


def main():
    with sync_playwright() as p:
        browser = p.firefox.launch(headless=False)
        page = browser.new_page()
        link = 'https://www.bseindia.com/stock-share-price/reliance-industries-ltd/reliance/500325/'
        page.goto(link, timeout=60000)
        time.sleep(2)
        l = link.split('/')
        company_name = l[4].strip().capitalize()

        df_corpgov_filtered = get_corpgov_data(page, company_name)
        df_equity = get_equity_data( page, company_name)
        df_financial_quarterly,df_financial_annual = financial_results(page,company_name)
        df_peer_quaterly,df_peer_annual,df_peer_dividend =get_peer_group(page,company_name)
        df_board_meetings = get_board_meetings(page,company_name)
        df_shareholder_meetings = get_shareholder_meetings(page,company_name)




    df_merged_meetings = df_board_meetings+df_shareholder_meetings

    for df in [df_equity, df_corpgov_filtered,df_peer_quaterly,df_peer_annual,df_peer_dividend , df_financial_quarterly,df_financial_annual,df_merged_meetings]:
        df['Company Name'] = company_name
        df['Industry Name'] = "Aerospace & Defense"


    excel_path = "Combined_data_BSE.xlsx"

    # Save all DataFrames to different sheets in an Excel file
    with pd.ExcelWriter(excel_path) as writer:
        df_peer_quaterly.to_excel(writer, sheet_name='Peer_Quarterly Data', index=False)
        df_peer_annual.to_excel(writer, sheet_name='Peer_Annual Data', index=False)
        df_peer_dividend.to_excel(writer, sheet_name='Peer_dividend Data', index=False)
        df_corpgov_filtered.to_excel(writer, sheet_name='Board_Members', index=False)
        df_equity.to_excel(writer, sheet_name='Equity Data', index=False)
        df_financial_quarterly.to_excel(writer, sheet_name='Financial_Quaterly Data', index=False)
        df_financial_annual.to_excel(writer, sheet_name='Financial_Annual Data', index=False)
        df_merged_meetings.to_excel(writer, sheet_name='Merged_Meetings Data', index=False)

if __name__ == "__main__":
    main()

from playwright.sync_api import sync_playwright
import time
from bs4 import BeautifulSoup
import re
import pandas as pd


# def get_equity_data(soup, page, company):

#     time.sleep(3)
#     dict1 = dict()

#     div_elements = soup.find_all('div', class_='col-lg-13')
#     for element in div_elements:
#         try:
#             tr_tags = element.find('tbody').find_all('tr')
#         except:
#             continue
#         for tr_tag in tr_tags:
#             try:
#                 key = tr_tag.find('td', class_='textsr').text.strip()
#                 value = tr_tag.find('td', class_='textvalue ng-binding').text.strip()
#                 dict1[key] = value
#             except:
#                 continue
    
#     row_list = []
#     data = dict1
#     row_list.append(data.values())
#     df = pd.DataFrame(row_list, columns=data.keys())
#     df.to_excel(f'{company}_equity_data.xlsx', index=False)


    
# def get_corpgov_data(soup, page, company):
#     page.get_by_role("link", name="Corporate Governance").click()
#     time.sleep(3)
#     element = soup.find_all('table', class_='ng-scope')[4]
#     row_list= list()
#     tr_body = element.find('tbody').find_all('tr')
#     # tr_heads = element.find('thead').find_all('tr')

#     for tr_tag in tr_body:
#         l = list()
#         try:
#             td_tags = tr_tag.find_all('td', class_='ng-binding')
#             for td_tag in td_tags:
#                 l.append(td_tag.text.strip())
            
#         except:
#             continue
#         row_list.append(l)
#     column_data = ['Sr','Title (Mr/Ms)','Name of the Director','DIN','Category','Whether the director is disqualified?','Start Date of disqualification','End Date of disqualification','Details of disqualification','Current status','Whether special resolution passed? [Refer Reg. 17(1A) of Listing Regulations]','Date of passing special resolution','Initial Date of Appointment','Date of Re-appointment','Date of cessation','Tenure of Director (in months)','No of Directorship in listed entities including this listed entity (Refer Regulation 17A of Listing Regulations)','No of Independent Directorship in listed entities including this listed entity (Refer Regulation 17A(1) of Listing Regulations','Number of memberships in Audit/ Stakeholder Committee(s) including this listed entity (Refer Regulation 26(1) of Listing Regulations)','No of post of Chairperson in Audit/ Stakeholder Committee held in listed entities including this listed entity (Refer Regulation 26(1) of Listing Regulations)','Reason for Cessation','Notes for not providing PAN','Notes for not providing DIN']
#     dataframe = pd.DataFrame(row_list,columns=column_data)
#     dataframe.to_excel(f'{company}_corp_gov.xlsx', index=False)
#     needed_row_list = list()
#     for lis in row_list:
#         row_list1 = [lis[0],lis[1],lis[2],lis[4],lis[9],lis[12],lis[13]]
#         needed_row_list.append(row_list1)
#     column_data1 = ['Sr','Title (Mr/Ms)','Name of the Director','Category','Current status','Initial Date of Appointment','Date of Re-appointment']
#     dataframe1 = pd.DataFrame(needed_row_list,columns=column_data1)
#     dataframe1.to_excel(f'{company}_corp_gov_needed.xlsx', index=False)



# def get_bulk_deals(page, company):
#     page.get_by_role("link", name="Bulk / Block deals").click()
#     time.sleep(5)
#     soup = BeautifulSoup(page.content(), 'html.parser')
#     element = soup.find('table', id='tblinsidertrd')
#     row_list= list()
#     tr_heads = element.find('thead').find_all('tr')
#     tr_body = element.find('tbody').find_all('tr')
#     columns = list()
#     for tr_tag in tr_body:
#         l = list()
#         try:
#             td_tags = tr_tag.find_all('td', class_='tdcolumn ng-binding')
#             for td_tag in td_tags:
#                 l.append(td_tag.text.strip())
#         except:
#             continue
#         row_list.append(l)
#     for th_tag in tr_heads:
#         td_tags = th_tag.find_all('td', class_='tableheading')
#         for td_tag in td_tags:
#             columns.append(td_tag.text.strip())
#     dataframe = pd.DataFrame(row_list,columns=columns)
#     dataframe.to_excel(f'{company}_bulk_deals.xlsx', index=False)



# def get_block_deals(page, company):
#     page.get_by_role("link", name="Bulk / Block deals").click()
#     time.sleep(3)
#     page.locator('input[name="dealtype"][value="2"][ng-model="Dealtype"]').click()
#     time.sleep(1)
#     page.locator('input[type="button"][value="Submit"]').click()
#     time.sleep(3)
#     soup = BeautifulSoup(page.content(), 'html.parser')
#     element = soup.find('table', id = 'tblinsidertrd')
#     row_list= list()
#     tr_heads = element.find('thead').find_all('tr')
#     tr_body = element.find('tbody').find_all('tr')
#     columns = list()
#     for tr_tag in tr_body:
#         l = list()
#         try:
#             td_tags = tr_tag.find_all('td', class_='tdcolumn ng-binding')
#             for td_tag in td_tags:
#                 l.append(td_tag.text.strip())
#         except:
#             continue
#         row_list.append(l)
#     for th_tag in tr_heads:
#         td_tags = th_tag.find_all('td', class_='tableheading')
#         for td_tag in td_tags:
#             columns.append(td_tag.text.strip())
#     dataframe = pd.DataFrame(row_list,columns=columns)
#     dataframe.to_excel(f'{company}_block_deals.xlsx', index=False)



# def get_financial_data(page, company):
    # def financial_results(page,company):
    #     page.locator("a#afi[data-toggle='collapse'][data-parent='#accordion']").click()
    #     time.sleep(2)
    #     page.locator("a#l61").click()
    #     time.sleep(3)
    #     soup = BeautifulSoup(page.content(), 'html.parser')
    #     element = soup.find_all('table', class_ ="ng-binding")[0]
    #     row_list= list()
    #     tr_heads = element.find('thead').find_all('tr')
    #     tr_body = element.find('tbody').find_all('tr')
    #     columns = list()
    #     for tr_tag in tr_body:
    #         l = list()
    #         try:
    #             td_tags = tr_tag.find_all('td', class_='tdcolumn')
    #             for td_tag in td_tags:
    #                 l.append(td_tag.text.strip())
    #         except:
    #             continue
    #         row_list.append(l)
    #     for th_tag in tr_heads:
    #         td_tags = th_tag.find_all('td', class_='tableheading')
    #         for td_tag in td_tags:
    #             columns.append(td_tag.text.strip())
    #     dataframe1 = pd.DataFrame(row_list,columns=columns)
    #     dataframe1 = dataframe1[:-3]
    #     dataframe1 =dataframe1.drop(dataframe1.index[0])
    #     page.get_by_role("link", name="Annual Trends").click()
    #     time.sleep(3)
    #     soup1 = BeautifulSoup(page.content(), 'html.parser')
    #     element1 = soup1.find_all('table', class_ ="ng-binding")[3]
    #     row_list1= list()
    #     tr_heads1 = element1.find('thead').find_all('tr')
    #     tr_body1 = element1.find('tbody').find_all('tr')
    #     columns1 = list()
    #     for tr_tag1 in tr_body1:
    #         l1 = list()
    #         try:
    #             td_tags1 = tr_tag1.find_all('td', class_='tdcolumn')
    #             for td_tag1 in td_tags1:
    #                 l1.append(td_tag1.text.strip())
    #         except:
    #             continue
    #         row_list1.append(l1)
    #     for th_tag1 in tr_heads1:
    #         td_tags1 = th_tag1.find_all('td', class_='tableheading')
    #         for td_tag1 in td_tags1:
    #             columns1.append(td_tag1.text.strip())
    #     dataframe2 = pd.DataFrame(row_list1,columns=columns1)
    #     dataframe2 = dataframe2[:-3]
    #     dataframe2 =dataframe2.drop(dataframe2.index[0])
    #     # dataframe = pd.concat([dataframe1,dataframe2], axis=1)
    #     dataframe1.to_excel(f'{company}_financial_results_quaterly.xlsx', index=False)
    #     dataframe2.to_excel(f'{company}_financial_results_annual.xlsx', index=False)




    # def financial_annual_reports(page,company):
    #     page.locator("a#afi[data-toggle='collapse'][data-parent='#accordion']").click()
    #     time.sleep(2)
    #     page.locator("a#162").click()
    #     time.sleep(3)
    #     soup = BeautifulSoup(page.content(), 'html.parser')
    #     element = soup.find_all('table', class_ ="ng-binding")[0]
    #     print(element)
    #     row_list= list()
    #     tr_heads = element.find('thead').find_all('tr')
    #     tr_body = element.find('tbody').find_all('tr')
    #     columns = list()
    #     for tr_tag in tr_body:
    #         l = list()
    #         try:
    #             td_tags = tr_tag.find_all('td', class_='tdcolumn')
    #             for td_tag in td_tags:
    #                 l.append(td_tag.text.strip())
    #         except:
    #             continue
    #         row_list.append(l)
    #     for th_tag in tr_heads:
    #         td_tags = th_tag.find_all('td', class_='tableheading')
    #         for td_tag in td_tags:
    #             columns.append(td_tag.text.strip())
    #     dataframe = pd.DataFrame(row_list,columns=columns)
    #     dataframe.to_excel(f'{company}_financial_results.xlsx', index=False)
    # financial_results(page,company)
    # financial_annual_reports(page,company)


# def get_meetings(page, company):
#     def get_board_meetings(page,company):
#         page.locator('a#ame').click()
#         time.sleep(3)
#         page.locator('a#l71').click()
#         time.sleep(6)
#         soup = BeautifulSoup(page.content(), 'html.parser')
#         element = soup.find_all('table', class_ = "ng-scope")[1]
#         row_list= list()
#         tr_heads = element.find('thead').find_all('tr')
#         tr_body = element.find('tbody').find_all('tr')
#         columns = list()
#         for tr_tag in tr_body:
#             l = list()
#             try:
#                 td_tags = tr_tag.find_all('td', class_='tdcolumn ng-binding')
#                 for td_tag in td_tags:
#                     l.append(td_tag.text.strip())
#             except:
#                 continue
#             row_list.append(l)
#         for th_tag in tr_heads:
#             td_tags = th_tag.find_all('td', class_='tableheading')
#             for td_tag in td_tags:
#                 columns.append(td_tag.text.strip())
#         dataframe = pd.DataFrame(row_list,columns=columns)
#         dataframe.to_excel(f'{company}_board_meetings.xlsx', index=False)
#     def get_shareholder_meetings(page, company):
#         # page.locator('a#ame').click()
#         # time.sleep(3)
#         page.locator('a#l72').click()
#         time.sleep(6)
#         soup = BeautifulSoup(page.content(), 'html.parser')
#         element = soup.find_all('table', class_ = "ng-scope")[1]
#         row_list= list()
#         tr_heads = element.find('thead').find_all('tr')
#         tr_body = element.find('tbody').find_all('tr')
#         columns = list()
#         for tr_tag in tr_body:
#             l = list()
#             try:
#                 td_tags = tr_tag.find_all('td', class_='tdcolumn ng-binding')
#                 for td_tag in td_tags:
#                     l.append(td_tag.text.strip())
#             except:
#                 continue
#             row_list.append(l)
#         for th_tag in tr_heads:
#             td_tags = th_tag.find_all('td', class_='tableheading')
#             for td_tag in td_tags:
#                 columns.append(td_tag.text.strip())
#         dataframe = pd.DataFrame(row_list,columns=columns)
#         dataframe.to_excel(f'{company}_shareholder_meetings.xlsx', index=False)
#     get_board_meetings(page,company)
#     get_shareholder_meetings(page, company)

# def get_corp_actions(page,company):
#     page.get_by_role("link", name="Corp Actions").click()
#     time.sleep(3)
#     soup = BeautifulSoup(page.content(), 'html.parser')
#     element = soup.find_all('table', class_ = "ng-scope")[1]
#     row_list= list()
#     tr_heads = element.find('thead').find_all('tr')
#     tr_body = element.find('tbody').find_all('tr')
#     columns = list()
#     for tr_tag in tr_body:
#         l = list()
#         try:
#             td_tags = tr_tag.find_all('td', class_='tdcolumn ng-binding')
#             for td_tag in td_tags:
#                 l.append(td_tag.text.strip())
#         except:
#             continue
#         row_list.append(l)
#     for th_tag in tr_heads:
#         td_tags = th_tag.find_all('td', class_='tableheading')
#         for td_tag in td_tags:
#             columns.append(td_tag.text.strip())
#     dataframe1 = pd.DataFrame(row_list,columns=columns)
#     element1 = soup.find_all('table', class_ = "ng-scope")[2]
#     row_list1= list()
#     tr_heads1 = element1.find('thead').find_all('tr')
#     tr_body1 = element1.find('tbody').find_all('tr')
#     columns1 = list()
#     for tr_tag1 in tr_body1:
#         l1 = list()
#         try:
#             td_tags1 = tr_tag1.find_all('td', class_='tdcolumn ng-binding')
#             for td_tag1 in td_tags1:
#                 l1.append(td_tag1.text.strip())
#         except:
#             continue
#         row_list1.append(l1)
#     for th_tag1 in tr_heads1:
#         td_tags1 = th_tag1.find_all('td', class_='tableheading')
#         for td_tag1 in td_tags1:
#             columns1.append(td_tag1.text.strip())
#     dataframe2 = pd.DataFrame(row_list1,columns=columns1)
#     element2 = soup.find_all('table', class_ = "ng-scope")[3]
#     row_list2= list()
#     tr_heads2 = element2.find('thead').find_all('tr')
#     tr_body2 = element2.find('tbody').find_all('tr')
#     columns2 = list()
#     for tr_tag2 in tr_body2:
#         l2 = list()
#         try:
#             td_tags2 = tr_tag2.find_all('td', class_='tdcolumn ng-binding')
#             for td_tag2 in td_tags2:
#                 l2.append(td_tag2.text.strip())
#         except:
#             continue
#         row_list2.append(l2)
#     for th_tag2 in tr_heads2:
#         td_tags2 = th_tag2.find_all('td', class_='tableheading')
#         for td_tag2 in td_tags2:
#             columns2.append(td_tag2.text.strip())
#     dataframe3 = pd.DataFrame(row_list2,columns=columns2)
#     with pd.ExcelWriter(f'{company}_corp_actions.xlsx', engine='xlsxwriter') as writer:
#     # Write the first DataFrame with its title
#         workbook = writer.book
#         worksheet = writer.sheets['Sheet1']
#         worksheet.write('A1', 'Corporate Actions')
#         worksheet.merge_range('A1:B1', 'Corporate Actions')  # Optional: Merge cells for the title
#         title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
#         worksheet.write('A1:B1', 'Corporate Actions', title_format)
#         worksheet.set_column('A:D', 18)

#         # Write the second DataFrame with its title
#         start_row = len(dataframe1) + 4  # Calculate starting row for the second DataFrame (after the first DataFrame + 1 for spacing)
#         dataframe2.to_excel(writer, sheet_name='Sheet1', startrow=start_row, index=False)
#         worksheet.write(start_row - 1, 0, 'Dividend Declared')  # Title for the second DataFrame
#         worksheet.merge_range(start_row - 1, 0, start_row - 1, 1, 'Dividend declared')  # Optional: Merge cells for the title
#         worksheet.write(start_row - 1, 0, 'Dividend declared', title_format)  # Title formatting
#         worksheet.set_column('A:D', 18) 

#         start_row1 = start_row + len(dataframe2) + 4  # Calculate starting row for the second DataFrame (after the first DataFrame + 1 for spacing)
#         dataframe3.to_excel(writer, sheet_name='Sheet1', startrow=start_row1, index=False)
#         worksheet.write(start_row1 - 1, 0, 'Bonus history')  # Title for the second DataFrame
#         worksheet.merge_range(start_row1 - 1, 0, start_row1 - 1, 1, 'Bonus history')  # Optional: Merge cells for the title
#         worksheet.write(start_row1 - 1, 0, 'Bonus history', title_format)  # Title formatting
#         worksheet.set_column('A:D', 18)



# def get_shareholding_pattern(page,company):
#     page.get_by_role("heading", name="Shareholding Pattern").get_by_role("link").click()
#     time.sleep(3)
#     soup = BeautifulSoup(page.content(), 'html.parser')
#     table = soup.find('table', class_ = "ng-binding")
#     element = table.find_all('table', border = "0")[2]
#     row_list= list()
#     tr_body = element.find('tbody').find_all('tr')
#     columns = list()
#     for tr_tag in tr_body:
#         l = list()
#         try:
#             td_tags = tr_tag.find_all('td', class_="TTRow_right")
#             for td_tag in td_tags:
#                 l.append(td_tag.text.strip())
#         except:
#             continue
#         row_list.append(l)
#     for th_tag in tr_body:
#         td_tags = th_tag.find_all('td', class_='innertable_header1')
#         for td_tag in td_tags:
#             columns.append(td_tag.text.strip())
#     del columns[10:12]
#     del columns[0]
#     del row_list[0:3]
#     del row_list[-1]
#     dataframe = pd.DataFrame(row_list,columns=columns)
#     dataframe.insert(0,'Category of shareholder',['(A) Promoter & Promoter Group','(B) Public','(C1) Shares underlying DRs','(C2) Shares held by Employee Trust','(C) Non Promoter-Non Public','Grand Total'])
#     dataframe.to_excel(f'{company}_Shareholding_pattern.xlsx', index=False)



# def get_investor_complaints(page,company):
#     page.get_by_role("heading", name="Statement of investor complaints").get_by_role("link").click()
#     time.sleep(3)
#     soup = BeautifulSoup(page.content(), 'html.parser')
#     element = soup.find_all('table', class_ = "ng-scope")[-1]
#     row_list= list()
#     tr_heads = element.find('thead').find_all('tr')
#     tr_body = element.find('tbody').find_all('tr')
#     columns = list()
#     for tr_tag in tr_body:
#         l = list()
#         try:
#             td_tags = tr_tag.find_all('td', class_='tdcolumn ng-binding')
#             for td_tag in td_tags:
#                 l.append(td_tag.text.strip())
#         except:
#             continue
#         row_list.append(l)
#     for th_tag in tr_heads:
#         td_tags = th_tag.find_all('td', class_='tableheading')
#         for td_tag in td_tags:
#             columns.append(td_tag.text.strip())
#     dataframe = pd.DataFrame(row_list,columns=columns)
#     dataframe.to_excel(f'{company}_investor_complaints.xlsx', index=False)


# def get_corp_information(page,company):
#     page.get_by_role("heading", name="Corp Information").get_by_role("link").click()
#     time.sleep(3)
#     soup = BeautifulSoup(page.content(), 'html.parser')
#     element = soup.find_all('table', class_ = "ng-scope")[1]
#     row_list= list()
#     tr_heads = element.find('thead').find_all('tr')
#     tr_body = element.find('tbody').find_all('tr')
#     for th_tag in tr_heads:
#         td_tags = th_tag.find_all('td', class_='blueheadertd')
#         for td_tag in td_tags:
#             heading = td_tag.text.strip()
#     for tr_tag in tr_body:
#         l = list()
#         try:
#             td_tags = tr_tag.find_all('td', class_='tdcolumn')
#             for td_tag in td_tags:
#                 l.append(td_tag.text.strip())
#         except:
#             continue
#         row_list.append(l)
#     dataframe = pd.DataFrame(row_list)
#     element1 = soup.find_all('table', class_ = "ng-scope")[2]
#     row_list1= list()
#     tr_heads1 = element1.find('thead').find_all('tr')
#     tr_body1 = element1.find('tbody').find_all('tr')
#     for th_tag1 in tr_heads1:
#         td_tags1 = th_tag1.find_all('td', class_='blueheadertd')
#         for td_tag1 in td_tags1:
#             heading1 = td_tag1.text.strip()
#     for tr_tag1 in tr_body1:
#         l1 = list()
#         try:
#             td_tags1 = tr_tag1.find_all('td', class_='tdcolumn')
#             for td_tag1 in td_tags1:
#                 l1.append(td_tag1.text.strip())
#         except:
#             continue
#         row_list1.append(l1)
#     dataframe1 = pd.DataFrame(row_list1)
#     element2 = soup.find_all('table', class_ = "ng-scope")[5]
#     row_list2= list()
#     tr_heads2 = element2.find('thead').find_all('tr')
#     tr_body2 = element2.find('tbody').find_all('tr')
#     for th_tag2 in tr_heads2:
#         td_tags2 = th_tag2.find_all('td', class_='blueheadertd')
#         for td_tag2 in td_tags2:
#             heading2 = td_tag2.text.strip()
#     for tr_tag2 in tr_body2:
#         l2 = list()
#         try:
#             td_tags2 = tr_tag2.find_all('td', class_='tdcolumn')
#             for td_tag2 in td_tags2:
#                 l2.append(td_tag2.text.strip())
#         except:
#             continue
#         row_list2.append(l2)
#     dataframe2 = pd.DataFrame(row_list2)
#     element3 = soup.find_all('table', class_ = "ng-scope")[7]
#     row_list3= list()
#     tr_heads3 = element3.find('thead').find_all('tr')
#     tr_body3 = element3.find_all('tbody')[1].find_all('tr')
#     print(tr_body3)
#     for th_tag3 in tr_heads3:
#         td_tags3 = th_tag3.find_all('td', class_='blueheadertd')
#         for td_tag3 in td_tags3:
#             heading3 = td_tag3.text.strip()
#     for tr_tag3 in tr_body3:
#         l3 = list()
#         try:
#             td_tags3 = tr_tag3.find_all('td', class_='tdcolumn')
#             for td_tag3 in td_tags3:
#                 l3.append(td_tag3.text.strip())
#         except:
#             continue
#         row_list3.append(l3)
#     dataframe3 = pd.DataFrame(row_list3)
#     element4 = soup.find_all('table', class_ = "ng-scope")[8]
#     row_list4= list()
#     tr_heads4 = element4.find('thead').find_all('tr')
#     tr_body4 = element4.find('tbody').find_all('tr')
#     for th_tag4 in tr_heads4:
#         td_tags4 = th_tag4.find_all('td', class_='blueheadertd')
#         for td_tag4 in td_tags4:
#             heading4 = td_tag4.text.strip()
#     for tr_tag4 in tr_body4:
#         l4 = list()
#         try:
#             td_tags4 = tr_tag4.find_all('td', class_='tdcolumn')
#             for td_tag4 in td_tags4:
#                 l4.append(td_tag4.text.strip())
#         except:
#             continue
#         row_list4.append(l4)
#     dataframe4 = pd.DataFrame(row_list4)
#     element5 = soup.find_all('table', class_ = "ng-scope")[9]
#     row_list5= list()
#     tr_heads5 = element5.find('thead').find_all('tr')
#     tr_body5 = element5.find('tbody').find_all('tr')
#     for th_tag5 in tr_heads5:
#         td_tags5 = th_tag5.find_all('td', class_='blueheadertd')
#         for td_tag5 in td_tags5:
#             heading5 = td_tag5.text.strip()
#     for tr_tag5 in tr_body5:
#         l5 = list()
#         try:
#             td_tags5 = tr_tag5.find_all('td', class_='tdcolumn')
#             for td_tag5 in td_tags5:
#                 l5.append(td_tag5.text.strip())
#         except:
#             continue
#         row_list5.append(l5)
#     dataframe5 = pd.DataFrame(row_list5)
#     with pd.ExcelWriter(f'{company}_corp_information.xlsx', engine='xlsxwriter') as writer:
#     # Write the first DataFrame with its title
#         dataframe.to_excel(writer, sheet_name='Sheet1', index=False)
#         workbook = writer.book
#         worksheet = writer.sheets['Sheet1']
#         worksheet.write('A1', 'Security details')
#         worksheet.merge_range('A1:B1', 'Security details')  # Optional: Merge cells for the title
#         title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
#         worksheet.write('A1:B1', 'Security details', title_format)
#         worksheet.set_column('A:D', 18)

#         # Write the second DataFrame with its title
#         start_row = len(dataframe) + 4  # Calculate starting row for the second DataFrame (after the first DataFrame + 1 for spacing)
#         dataframe1.to_excel(writer, sheet_name='Sheet1', startrow=start_row, index=False)
#         worksheet.write(start_row - 1, 0, 'Board of Directors')  # Title for the second DataFrame
#         worksheet.merge_range(start_row - 1, 0, start_row - 1, 1, 'Board of Directors')  # Optional: Merge cells for the title
#         worksheet.write(start_row - 1, 0, 'Board of Directors', title_format)  # Title formatting
#         worksheet.set_column('A:D', 18) 

#         start_row1 = start_row + len(dataframe1) + 4  # Calculate starting row for the second DataFrame (after the first DataFrame + 1 for spacing)
#         dataframe2.to_excel(writer, sheet_name='Sheet1', startrow=start_row1, index=False)
#         worksheet.write(start_row1 - 1, 0, 'Statutory Auditor Details')  # Title for the second DataFrame
#         worksheet.merge_range(start_row1 - 1, 0, start_row1 - 1, 1, 'Statutory Auditor Details')  # Optional: Merge cells for the title
#         worksheet.write(start_row1 - 1, 0, 'Statutory Auditor Details', title_format)  # Title formatting
#         worksheet.set_column('A:D', 18)

#         start_row2 = start_row1 + len(dataframe2) + 4  # Calculate starting row for the second DataFrame (after the first DataFrame + 1 for spacing)
#         dataframe3.to_excel(writer, sheet_name='Sheet1', startrow=start_row2, index=False)
#         worksheet.write(start_row2 - 1, 0, 'Secretarial Auditor details')  # Title for the second DataFrame
#         worksheet.merge_range(start_row2 - 1, 0, start_row2 - 1, 1, 'Secretarial Auditor details')  # Optional: Merge cells for the title
#         worksheet.write(start_row2 - 1, 0, 'Secretarial Auditor details', title_format)  # Title formatting
#         worksheet.set_column('A:D', 18)

#         start_row3 = start_row2 + len(dataframe3) + 4  # Calculate starting row for the second DataFrame (after the first DataFrame + 1 for spacing)
#         dataframe4.to_excel(writer, sheet_name='Sheet1', startrow=start_row3, index=False)
#         worksheet.write(start_row3 - 1, 0, 'Registered Office Address')  # Title for the second DataFrame
#         worksheet.merge_range(start_row3 - 1, 0, start_row3 - 1, 1, 'Registered Office Address')  # Optional: Merge cells for the title
#         worksheet.write(start_row3 - 1, 0, 'Registered Office Address', title_format)  # Title formatting
#         worksheet.set_column('A:D', 18)

#         start_row4 = start_row3 + len(dataframe4) + 4  # Calculate starting row for the second DataFrame (after the first DataFrame + 1 for spacing)
#         dataframe5.to_excel(writer, sheet_name='Sheet1', startrow=start_row4, index=False)
#         worksheet.write(start_row4 - 1, 0, 'Registrars')  # Title for the second DataFrame
#         worksheet.merge_range(start_row4 - 1, 0, start_row4 - 1, 1, 'Registrars')  # Optional: Merge cells for the title
#         worksheet.write(start_row4 - 1, 0, 'Registrars', title_format)  # Title formatting
#         worksheet.set_column('A:D', 18)


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
    tr_heads2 = element8.find('thead').find_all('tr')
    tr_body2 = element8.find('tbody').find_all('tr')
    columns2 = list()
    for tr_tag2 in tr_body2:
        l2 = list()
        try:
            td_tags2 = tr_tag2.find_all('td', class_='tdcolumn')
            for td_tag2 in td_tags2:
                l2.append(td_tag2.text.strip())
        except:
            continue
        row_list2.append(l2)
    for th_tag2 in tr_heads2:
        td_tags2 = th_tag2.find_all('td', class_='tableheading')
        for td_tag2 in td_tags2:
            columns2.append(td_tag2.text.strip())
    dataframe3 = pd.DataFrame(row_list2,columns=columns2)
    # # dataframe = pd.concat([dataframe1,dataframe2], axis=1)
    file_path = f'{company}_peer_group.xlsx'

# Use ExcelWriter to write multiple DataFrames to different sheets
    with pd.ExcelWriter(file_path) as writer:
        dataframe1.to_excel(writer, sheet_name='Sheet1', index=False)
        dataframe2.to_excel(writer, sheet_name='Sheet2', index=False)
        dataframe3.to_excel(writer, sheet_name='Sheet3', index=False)

    # dataframe1.to_excel(f'{company}_financial_results_quaterly.xlsx', index=False)
    # dataframe2.to_excel(f'{company}_financial_results_annual.xlsx', index=False)


                                                

def main():
    with sync_playwright() as p:
        browser = p.firefox.launch(headless=False)
        page = browser.new_page()
        link = 'https://www.bseindia.com/stock-share-price/reliance-industries-ltd/reliance/500325/'
        page.goto(link, timeout=60000)
        time.sleep(2)
        l = link.split('/')
        company = l[4].strip().capitalize()
        # get_corpgov_data(soup, page, company)
        # get_equity_data(soup, page, company)
        # get_bulk_deals(soup, page, company)
        # get_block_deals(page, company)
        # get_financial_data(page, company)
        # get_meetings(page, company)
        # get_corp_actions(page,company)
        # get_shareholding_pattern(page,company)
        # get_investor_complaints(page,company)
        # get_corp_information(page,company)
        get_peer_group(page,company)


if __name__ == "__main__":
    main()
from selenium import webdriver
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import pandas as pd

def workdays(d, end, excluded=(6, 7)):
    days = []
    while d.date() <= end.date():
        if d.isoweekday() not in excluded:
            days.append(d)
        d += timedelta(days=1)
    return days

stock_code = input('Input stock code: ')

start_date = '2020/01/15'
end_date = '2020/01/22'

start_date = datetime.strptime(start_date, '%Y/%m/%d')
end_date = datetime.strptime(end_date, '%Y/%m/%d')

working_days = workdays(start_date, end_date)
working_days = [datetime.strftime(i, '%Y/%m/%d') for i in working_days]

driver = webdriver.Chrome('chromedriver.exe')   
URL = 'http://www.hkexnews.hk/sdw/search/searchsdw.aspx'
driver.get(URL)

list_of_participants_id = []
list_of_participants = []
list_of_shareholding = []
list_of_dates = []

for j in working_days:
    # click stock code 
    stock_code_field = driver.find_elements_by_xpath("//input[@name='txtStockCode']")[0]
    stock_code_field.click()

    # input stock code
    stock_code_field.send_keys(stock_code)

    # input date
    select = driver.find_element_by_id("txtShareholdingDate")
    driver.execute_script("arguments[0].value = arguments[1]", select, j)

    #click search button
    search_button = driver.find_elements_by_xpath("//a[@class='btn-blue']")[0]
    search_button.click()

    subhtml = driver.page_source
    soup = BeautifulSoup(subhtml, "html.parser")
    
    for i in soup.find_all(class_="col-participant-id")[1:]:
        list_of_participants_id.append(i.find(class_="mobile-list-body").text)
    
    for i in soup.find_all(class_="col-participant-name")[1:]:
        list_of_participants.append(i.find(class_="mobile-list-body").text)
        list_of_dates.append(j)
        
    for i in soup.find_all(class_="col-shareholding")[1:]:
        list_of_shareholding.append(i.find(class_="mobile-list-body").text)

d = {'Date': list_of_dates, 'Participant ID':list_of_participants_id,'Participant':list_of_participants, 'Shareholding': list_of_shareholding}

df = pd.DataFrame(data=d)

choice_of_output = input("Type 'top 20' to get the Top 20 participant IDs or the participant ID for just one: ")

if choice_of_output[0:3] == 'top':
    input_part_ids = []

    for i in df['Participant ID'].head(20):
        input_part_ids.append(i)
    
    result_df = pd.DataFrame({'Date':[],'Participant ID': [], 'Participant': [], 'Shareholding': []})
    for i in input_part_ids:
        df1 = df[df['Participant ID']==i]
        # Convert string to number shareholding
        a = []
        for i in df1['Shareholding']:
            a.append(int(i.replace(',','')))

        df1.update(pd.DataFrame({'Shareholding': a}, index = df1.index))
        # Convert string to datetime
        b = []
        for i in df1['Date']:
            b.append(datetime.strptime(i, '%Y/%m/%d'))
        df1 = df1.astype({'Shareholding': 'int32'})
        result_df = result_df.append(df1)
        
    result_df.to_excel('output.xlsx')    

else:
    df1 = df[df['Participant ID']==choice_of_output]

    a = []
    for i in df1['Shareholding']:
        a.append(int(i.replace(',','')))

    df1.update(pd.DataFrame({'Shareholding': a}, index = df1.index))

    b = []
    for i in df1['Date']:
        b.append(datetime.strptime(i, '%Y/%m/%d'))
    
    df1 = df1.astype({'Shareholding': 'int32'})

    df1.to_excel('output.xlsx')

driver.close()
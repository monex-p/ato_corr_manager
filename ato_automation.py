from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
# from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep
from datetime import datetime, timedelta

import pandas as pd
import re
import os

import xlwings as xw

# options = Options()
# options.headless = True
# options.add_argument("--window-size=1920,1200")

# get excel file
excel_file_name = "ATO Correspondence Master 2023.xlsm"
yesterday = datetime.today() - timedelta(days=1)
sheet_name = yesterday.strftime("%b%Y")
# sheet_name = "Apr2023"
parent_dir = os.path.dirname(os.getcwd())
excel_path = os.path.join(parent_dir, excel_file_name)

wb = xw.Book(excel_path)
sheet = wb.sheets[sheet_name]

col_index = 6 # row of issue date
last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
start_date = datetime.strptime(sheet.range((last_row,col_index)).value, "%d/%m/%Y") + timedelta(days=1)

### how to deal with different months?

ato_website = 'https://onlineservices.ato.gov.au/OnlineServices/home#home'

def ato_login(website):
    driver = webdriver.Chrome()
    # open browser session
    driver.get(website)
    ### may want to auto fill log in email
    driver.implicitly_wait(60)

comm_col_id = "atoo-ahp-atomastermenu-001-3"
comm_hist_id = "atoo-ahp-atomastermenu-001-3-1"
time_arr_id = "dd-atoo-cch-time-period-001"
choose_date_xpath = "//option[@value='CD']"
start_date_id = "dp-atoo-cch-from-001"
end_date_id = "dp-atoo-cch-to-001"
email_box_id = "cbl--atoo-cch-channel-002-1atoo-cch-email-001"
sms_box_id = "cbl--atoo-cch-channel-002-2atoo-cch-sms-001"
button_id = "atoo-cch-atobutton-012"

def request_corr(comm_col_id, comm_hist_id, time_arr_id, choose_date_xpath,
                 start_date_id, end_date_id, email_box_id, sms_box_id, button_id):
    # choose communication history tab
    driver.find_element(By.ID, comm_col_id).click()
    driver.find_element(By.ID, comm_hist_id).click()

    # choose date range in consideration
    driver.find_element(By.ID, time_arr_id).click()
    driver.find_element(By.XPATH, choose_date_xpath).click()

    driver.find_element(By.ID, start_date_id).send_keys(start_date.strftime("%d/%m/%Y"))
    # driver.find_element(By.ID, "dp-atoo-cch-from-001").send_keys("01/05/2023")
    driver.find_element(By.ID, end_date_id).send_keys(yesterday.strftime("%d/%m/%Y"))
    # driver.find_element(By.ID, "dp-atoo-cch-to-001").send_keys("11/05/2023")

    # uncheck email and sms corr
    email_checkbox = driver.find_element(By.ID, email_box_id)
    sms_checkbox = driver.find_element(By.ID, sms_box_id)
    driver.execute_script("arguments[0].click();", email_checkbox)
    driver.execute_script("arguments[0].click();", sms_checkbox)

    driver.find_element(By.ID, button_id).click() # submit form


def corr_table_interact():
    # change capacity to 100/page
    driver.find_element(By.ID, "dd-atoo-cch-results-per-page-001").click()
    driver.find_element(By.XPATH, "//option[@value='100']").click()

    ### think of how to deal w more than 100 corres per time

    # getting table info
    sleep(2) # waiting table to load
    ato_table = driver.find_element(By.ID, "atoo-cch-ato-table-001")
    string_list = ato_table.text.split("\n")
    chunks = [string_list[i:i+5] for i in range(0, len(string_list), 5)]
    chunks = chunks[1:][::-1]

    ato_df = pd.DataFrame(chunks, columns=['Name', 'Client ID', 'Subject', 'Channel', 'Issue Date'])

    # download corr and get corr ID list - need to prepare a list of subjects not gonna download/ consider
    corr_id_list = []
    sleep(1.8)
    for i in range(len(chunks)):
        element = driver.find_element(By.ID, "atoo-cch-atolink-corres-%s"%(i))
        driver.execute_script("arguments[0].click();", element)
        sleep(2)
        href = element.get_attribute("href")
        start_ind = href.find("ID=") + 3
        end_ind = href.find("#Corr")
        corr_id = href[start_ind:end_ind]
        corr_id_list.append(corr_id)

    ato_df["Correspondence"] = corr_id_list[::-1]

    return ato_df

##################################################################

driver = webdriver.Chrome()

# open browser session
driver.get('https://onlineservices.ato.gov.au/OnlineServices/home#home')
driver.implicitly_wait(60)

# choose communication history tab
driver.find_element(By.ID, "atoo-ahp-atomastermenu-001-3").click()
driver.find_element(By.ID, "atoo-ahp-atomastermenu-001-3-1").click()

# choose date range in consideration
driver.find_element(By.ID, "dd-atoo-cch-time-period-001").click()
driver.find_element(By.XPATH, "//option[@value='CD']").click()

driver.find_element(By.ID, "dp-atoo-cch-from-001").send_keys(start_date.strftime("%d/%m/%Y"))
# driver.find_element(By.ID, "dp-atoo-cch-from-001").send_keys("01/05/2023")
driver.find_element(By.ID, "dp-atoo-cch-to-001").send_keys(yesterday.strftime("%d/%m/%Y"))
# driver.find_element(By.ID, "dp-atoo-cch-to-001").send_keys("11/05/2023")

# uncheck email and sms corr
email_checkbox = driver.find_element(By.ID, "cbl--atoo-cch-channel-002-1atoo-cch-email-001")
sms_checkbox = driver.find_element(By.ID, "cbl--atoo-cch-channel-002-2atoo-cch-sms-001")
driver.execute_script("arguments[0].click();", email_checkbox)
driver.execute_script("arguments[0].click();", sms_checkbox)

driver.find_element(By.ID, "atoo-cch-atobutton-012").click() # submit form

# change capacity to 100/page
driver.find_element(By.ID, "dd-atoo-cch-results-per-page-001").click()
driver.find_element(By.XPATH, "//option[@value='100']").click()

### think of how to deal w more than 100 corres per time

# getting table info
sleep(2) # waiting table to load
ato_table = driver.find_element(By.ID, "atoo-cch-ato-table-001")
string_list = ato_table.text.split("\n")
chunks = [string_list[i:i+5] for i in range(0, len(string_list), 5)]
chunks = chunks[1:][::-1]

ato_df = pd.DataFrame(chunks, columns=['Name', 'Client ID', 'Subject', 'Channel', 'Issue Date'])

# download corr and get corr ID list - need to prepare a list of subjects not gonna download/ consider
corr_id_list = []
sleep(1.8)
for i in range(len(chunks)):
    element = driver.find_element(By.ID, "atoo-cch-atolink-corres-%s"%(i))
    driver.execute_script("arguments[0].click();", element)
    sleep(2)
    href = element.get_attribute("href")
    start_ind = href.find("ID=") + 3
    end_ind = href.find("#Corr")
    corr_id = href[start_ind:end_ind]
    corr_id_list.append(corr_id)

ato_df["Correspondence"] = corr_id_list[::-1]

# change dataframe index to match excel index
ato_df.index = ato_df.index + last_row

# add the DataFrame to Excel starting at the next available row
sheet.range('A' + str(last_row + 1)).options(header=False).value = ato_df

# paste formula concurrently to new rows
copy_range = sheet.range("H{0}:L{0}".format(str(last_row)))
paste_range = sheet.range("H%s:L%s" % (str(last_row),str(ato_df.index[-1]+1)))
paste_range.formula = copy_range.formula



wb.save()    

print("**Finish Successfully**")
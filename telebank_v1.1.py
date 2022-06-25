import os
import sys
import time
import shutil
import urllib
import warnings
import traceback
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementNotInteractableException, StaleElementReferenceException, NoSuchElementException, WebDriverException, ElementClickInterceptedException, TimeoutException

import utils
import config
from helpers.g_sheet_handler import GoogleSheetHandler

warnings.filterwarnings("ignore")

class DataScrapping():
    
    def __init__(
                self, browser, username, password, recent_transactions, dep_checks, letters, check_date_diff, transaction_date_diff,
                start_date_transactions, start_date_checks, start_date_letters, gdrive_folder_id
         ):
        self.browser = browser
        self.username = username
        self.password = password
        self.user_login = False
        self.users = users
        self.recent_transactions = recent_transactions
        self.dep_checks = dep_checks
        self.letters = letters
        self.start_date_transactions = start_date_transactions
        self.start_date_checks = start_date_checks
        self.start_date_letters = start_date_letters
        self.gdrive_folder_id = gdrive_folder_id
        self.check_date_diff = check_date_diff
        self.transaction_date_diff = transaction_date_diff

        self.dep_checks_data = []
        self.recent_transactions_data = []
        self.image_name = {}
        self.all_data_map = {'dep_checks_data': config.SHEET_CHECKS_TAB_NAME+'_'+self.username, 'recent_transactions_data': config.SHEET_DATA_TAB_NAME+'_'+self.username}

    def login_to_site(self):
        print("\n\tStart user login..\n")
        try:
            self.browser.get(config.WEB_LINK)
            time.sleep(3)
            self.browser.find_element(By.ID, 'tzId').send_keys(self.username)
            self.browser.find_element(By.ID, 'tzPassword').send_keys(self.password)
            time.sleep(3)
            self.browser.find_element(By.CSS_SELECTOR, 'button[type="submit"]').click()
            self.user_login = True
            time.sleep(3)
            if self.browser.current_url == 'https://start.telebank.co.il/login/GENERAL_ERROR':
                print('\n\tRetrying... Login!\n')
                self.browser.delete_all_cookies()
                self.login_to_site()
            print('\n\tLogin successfull\n')
        except:
            print("\n\tLogin Failed !!\n")
            time.sleep(3)
            self.login_to_site()

    def create_directories(self):
        print(f'\n\n\t\t* * * * ** Creating Directory for Images & PDF for USER: {self.username} * * * * * * ')
        utils.create_dir(f'images_{self.username}')
        utils.create_dir('pdf')
        if os.path.isdir(f'csv_{self.username}'):
            print('\n\tCSV directory already exists!\n')
        else:
            os.makedirs(f'csv_{self.username}')
        
    def logout(self):
        time.sleep(5)
        try:
            self.browser.find_element(By.CSS_SELECTOR, '#logOutLink').click()
            self.user_login = False
            time.sleep(5)
            print('\n\tLogged out user(%s) successfully!\n' %self.username)
            return self.user_login
        except ElementNotInteractableException:
            print('ElementNotInteractableException Occured: [IGNORE]')
            pass
        except NoSuchElementException:
            print('NoSuchElementException Occured: [IGNORE]')
            pass
        except StaleElementReferenceException:
            print('StaleElementReferenceException Occured: [IGNORE]')
            pass
        except ElementClickInterceptedException:
            print('ElementClickInterceptedException Occured: [IGNORE]')
            pass
  
    def get_dep_checks_data(self):
        if self.dep_checks.upper() == 'NO':
            print('\n\tSkipping Deposit Checks Data\n')
            return False
        date, ref = utils.get_last_sheet_record(self.all_data_map['dep_checks_data'], self.username)
        # date = date[6:] + '/' + date[4:6] + '/' + date[:4]
        print(f'\n\tDATE:{date}, REF:{ref}\n')
        print('\n\tGetting Deposit Checks Data....\n')
        time.sleep(3)
        self.browser.get('https://start.telebank.co.il/apollo/business/#/CHKVEW')
        self.browser.refresh()
        time.sleep(7)
        if self.check_date_diff >= 360:
            WebDriverWait(self.browser,10).until(EC.element_to_be_clickable((By.ID,"dropdownMenu2"))).click()
            WebDriverWait(self.browser,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="checks-by-dates-4"]/a'))).click()
        else:
            self.filter_data_for_dep_checks()
        time.sleep(3)
        df = pd.DataFrame()
        try:
            df = utils.get_table_df(self.browser.page_source, 'sortable-table')
            df.insert(0, 'תאריך ריצה', datetime.now().strftime("%d/%m/%Y"))
            del df['Unnamed: 8']
            df['תאריך הפקדה'] = df['תאריך הפקדה'].str.replace('לחץ למידע נוסף על שורה זאת בטבלה', '')
            if date and ref:
                date = pd.to_datetime(date, format='%d/%-m/%Y', infer_datetime_format=True)
                df['תאריך הפקדה'] = pd.to_datetime(df['תאריך הפקדה'], format='%d/%-m/%Y', infer_datetime_format=True)
                mask = ((df['תאריך הפקדה'] >= date) & (df['סניף נמשך'] > int(ref)) | (df['תאריך הפקדה'] > date) & (df['סניף נמשך'] >= int(ref)))
                df = df.loc[mask]
                df['תאריך הפקדה'] = df['תאריך הפקדה'].astype(str)
                if df.empty:
                    mask = ((df['תאריך הפקדה'] <= date) & (df['סניף נמשך'] < int(ref)) | (df['תאריך הפקדה'] < date) & (df['סניף נמשך'] <= int(ref)))
                    df = df.loc[mask]
                    df['תאריך הפקדה'] = df['תאריך הפקדה'].astype(str)
            if not df.empty:
                front_img_li, back_img_li = self.get_check_image(df.values.tolist())
                df['קישור לתמונה קדמית'] = front_img_li
                df['קישור לתמונה אחורה'] = back_img_li
            print(df)
            data = df.values.tolist()
            self.dep_checks_data = data
            if df.empty:
                return False
            return True

        except WebDriverException:
            if df.empty:
                return False
            print("\n\tCouldn't find the page [RETRYING]\n")
            self.browser.refresh()
            time.sleep(15)
            self.get_dep_checks_data()  
        
        except Exception as err:
            print(f"Error in : {traceback.format_exc()}")
            print("\t\nNo table found between selected dates!\n")
            return False
    
    def get_check_image(self, data):
        data_li = [row[5:7] for row in data]
        front_img_li = []
        back_img_li = []
        row = 0
        try: 
            data_records = self.browser.find_elements(By.XPATH, '//*[@id="main-content"]/div[2]/feature-flag/div/flag-off-component/div/div/div/div[2]/div[1]/div[2]/div/div/section/table/tbody/tr')
            print(f'\n\tTotal records: {len(data_records)}\n')
            for row_idx in range(len(data_records)):
                time.sleep(3)
                # tr = self.browser.find_elements(By.XPATH, f'//*[@id="main-content"]/div[2]/feature-flag/div/flag-off-component/div/div/div/div[2]/div[1]/div[2]/div/div/section/table/tbody/tr[{row_idx+1}]')
                tr = WebDriverWait(self.browser, 20).until(EC.visibility_of_all_elements_located((By.XPATH, f'//*[@id="main-content"]/div[2]/feature-flag/div/flag-off-component/div/div/div/div[2]/div[1]/div[2]/div/div/section/table/tbody/tr[{row_idx+1}]')))
                for data in tr:
                    data_row = data.text.split(' ')
                check_number = data_row[10]
                print('\n\tProcessing Data for Check No: ', check_number)
                print(data_row)
                WebDriverWait(self.browser, 15).until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="main-content"]/div[2]/feature-flag/div/flag-off-component/div/div/div/div[2]/div[1]/div[2]/div/div/section/table/tbody/tr[{row_idx+1}]'))).click()
                if len(data_row) != 14:
                    front_img_li.append('Failed')
                    back_img_li.append('Failed')
                    continue
                time.sleep(3)
                front_image = self.browser.find_element(By.CSS_SELECTOR, 'body > div.modal.discountBiz-modal-general.cs-spa-sme-content.topbar-modal.checksExpand.fade.ng-scope.ng-isolate-scope.ng-animate.ng-enter.ng-enter-active.in > div > div > div > div.containerOsh.container-fluid > div.contentOsh.row > section.col-sm-6.col-print-9 > div > div > img').get_attribute('src')
                WebDriverWait(self.browser,10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'body > div.modal.discountBiz-modal-general.cs-spa-sme-content.topbar-modal.checksExpand.fade.ng-scope.ng-isolate-scope.ng-animate.ng-enter.ng-enter-active.in > div > div > div > div.containerOsh.container-fluid > div.contentOsh.row > section.col-sm-6.col-print-9 > button > img'))).click()
                time.sleep(2)
                back_image = self.browser.find_element(By.CSS_SELECTOR, 'body > div.modal.discountBiz-modal-general.cs-spa-sme-content.topbar-modal.checksExpand.fade.ng-scope.ng-isolate-scope.ng-animate.ng-enter.ng-enter-active.in > div > div > div > div.containerOsh.container-fluid > div.contentOsh.row > section.col-sm-6.col-print-9 > div > div > img').get_attribute('src')
                front_img_name = f'images_{self.username}/{check_number.replace("/", "-")}_1.jpg'
                urllib.request.urlretrieve(front_image, front_img_name)
                front_img_li.append(front_img_name)
                back_img_name = f'images_{self.username}/{check_number.replace("/", "-")}_2.jpg'
                urllib.request.urlretrieve(back_image, back_img_name)
                back_img_li.append(back_img_name)
                time.sleep(2)
                self.browser.find_element(By.CSS_SELECTOR, 'button[type="button"]').click()
                row += 1
            return front_img_li, back_img_li
        except WebDriverException:
            print("\n\tCouldn't find the page [RETRYING]\n")
            self.browser.refresh()
            time.sleep(15)
            self.get_check_image()

    def get_reference_index(self, ref, data):
        print('\n\tGetting last Saved Index\n')
        count = 0
        try:
            while True:
                if data == 'Checks':
                    refs = self.browser.find_element(By.XPATH, f'//*[@id="main-content"]/div[2]/feature-flag/div/flag-off-component/div/div/div/div[2]/div[1]/div[2]/div/div/section/table/tbody/tr[{count+1}]/td[3]')
                if data == 'Transactions':
                    refs = self.browser.find_element(By.XPATH, f'//*[@id="lastTransactionTable-cell-{count}-3"]')
                count+=1
                if refs.text == ref:
                    break
        except NoSuchElementException or WebDriverException as err:
            print(f"{err} Occured!")
            count = 0
        return count

    def get_recent_transaction_data(self):
        if self.recent_transactions.upper() == 'NO':
            print('\n\tSkipping Transactions Data\n')
            return
        refs_from_g_drive, date = utils.get_last_sheet_record(self.all_data_map['recent_transactions_data'], self.username)
        print('\n\tGetting Image Data from Recent Transactions\n')
        time.sleep(2)
        self.browser.get('https://start.telebank.co.il/apollo/business/#/OSH_LENTRIES_ALTAMIRA')
        self.browser.refresh()
        time.sleep(6)
        self.browser.implicitly_wait(20)
        if self.transaction_date_diff >= 360:
            WebDriverWait(self.browser,20).until(EC.element_to_be_clickable((By.ID,"input-osh-transaction"))).click()
            WebDriverWait(self.browser,20).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="lobby-osh-filter-item-7"]/a'))).click()
        else:
            self.filter_data_for_recent_transactions()
        time.sleep(7)
        self.browser.implicitly_wait(20)
        last_row_idx = 0
        if refs_from_g_drive:
            last_row_idx = int(self.get_reference_index(refs_from_g_drive[-1], 'Transactions'))
        else:
            refs_from_g_drive = [0]
        print('Last saved Index:', last_row_idx)
        print('Please Wait! Skipping to the Last Saved Index......')
        try:
            print('\n\t Getting Data\n')
            time.sleep(3)
            idx = 0
            data_records = self.browser.find_elements(By.CSS_SELECTOR, '.rc-table-row-content') 
            print('Total Records:',len(data_records))
            while idx < len(data_records):
                for data_row in data_records:
                    if idx < last_row_idx:
                        idx += 1
                        continue
                    if (not data_row) or (not data_row.text.strip()):
                        print(f'ids:{idx} - data_row not found hance trying again to cover missing one!')
                        self.search_for_date_range_again()
                        break
                    
                    row_date, row_ref = data_row.text.split('\n')[0], data_row.text.split('\n')[3]
                    check_no = utils.get_check_no(data_row.text.split()[3])
                    data_row_text = data_row.text.split('\n')
                    print("\tProcess row data: ref:", data_row.text.split('\n')[3])

                    if row_ref.split('/')[0] not in refs_from_g_drive or int(row_ref.split('/')[0]) > int(refs_from_g_drive[-1]):
                        WebDriverWait(self.browser,20).until(EC.element_to_be_clickable((By.ID,f"lastTransactionTable-row-{idx}"))).click()
                        time.sleep(5)
                        self.browser.implicitly_wait(30)
                        self.unique_reference = WebDriverWait(self.browser,10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, '.text-box-value-color.ng-binding.number-font'))).text
                        img_element = utils.verify_element(browser=self.browser, by_selector=By.XPATH, path='.//*[@id="single-check-view-con"]/div/div[1]/img[1]')
                        self.image_name[self.unique_reference] = ['', '', '']
                        time.sleep(3)
                        try:
                            self.browser.find_element(By.XPATH, '//*[@id="expanded-view-popup"]/div[4]/div[1]/div/div/div/ul/li[2]/a').click()
                            print('\tPDF Downloading')
                        except NoSuchElementException:
                            print('\t No PDF Found!')
                            pass
                        if img_element:
                            print('\tImg found!')
                            time.sleep(15)  
                            self.browser.implicitly_wait(17)
                            pdfs = os.listdir('pdf/tmp')
                            pdf_ref = row_ref.split('/')[0]
                            new_pdf_name = f'{pdf_ref}_{self.username}_{check_no}.pdf'
                            print(f"idx:{idx}, last_row_idx: {last_row_idx}")
                            utils.rename_file('pdf/tmp', pdfs[0], new_pdf_name) #Renaming pdf file 
                            shutil.move('pdf/tmp/'+new_pdf_name, "pdf")

                            self.download_image(check_no)
                            self.get_image_table_data(data_row_text, check_no, new_pdf_name)
                        else:
                            self.get_no_image_table_data(data_row)
                        try:
                            self.browser.find_element(By.CSS_SELECTOR, 'button[type="button"]').click();time.sleep(1.5)
                        except ElementNotInteractableException or StaleElementReferenceException:
                            print('Not Interactable Image: [IGNORE]')
                            pass
                        idx += 1

                    else:
                        print(f'Data Already Exists')
                        idx += 1
                        if idx+1 == len(data_records):
                            break
                        continue
            flag = utils.verify_element(browser=browser, by_selector=By.XPATH, path=f'//*[@id="lastTransactionTable-cell-{idx}-0"]')
            
        except WebDriverException:
            flag = utils.verify_element(browser=browser, by_selector=By.XPATH, path=f'//*[@id="lastTransactionTable-cell-{idx}-0"]')
            print(f"\t\n Error Occured Due to: {WebDriverException} ")
            print("\t\n\n ** ** ** ** ** Uploading the extracted imgs & pdf to gdrive and Pushing data to sheet till now  ** ** ** **")

        return flag

    def download_image(self, check_no):
        front_image = WebDriverWait(self.browser,10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="single-check-view-con"]/div/div[1]/img[1]'))
        ).get_attribute('src') 
        front_img_name = f'images_{self.username}/{self.unique_reference.replace("/", "-")}'+ "_"+ self.username +'_'+check_no + "_1.jpg"
        urllib.request.urlretrieve(front_image, front_img_name)
        
        back_image = WebDriverWait(self.browser,20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="single-check-view-con"]/div/div[1]/img[2]'))
        ).get_attribute('src')
        back_img_name = f'images_{self.username}/{self.unique_reference.replace("/", "-")}'+ "_"+ self.username +'_'+check_no +"_2.jpg"
        urllib.request.urlretrieve(back_image, back_img_name)
        return front_img_name, back_img_name

    def get_image_table_data(self, data_row, check_no, pdf_name):
        # print(f"====================data_row={data_row} \n\n check_no={check_no}")
        time.sleep(3)
        channel_name = self.browser.find_element(By.XPATH, '//*[@id="expanded-view-popup"]/div[4]/div[2]/div[2]/div/div[1]/div/ng-include/div/div[2]/span[2]/span[2]').text
        all_details = utils.get_check_table_df(self.browser.page_source, 'מספר בנק')
        front_img_name, back_img_name = self.download_image(check_no)
        self.image_name[self.unique_reference] = [front_img_name, back_img_name, pdf_name]
        self.recent_transactions_data.append([datetime.now().strftime("%d/%m/%Y")] + data_row + [channel_name]+ all_details.iloc[:][1].values.tolist() +["" , "", ""])
        print(self.recent_transactions_data[-1])
        print("\n")

    def get_no_image_table_data(self, data_row):
        try:
            comment = ''
            try:
                channel_name = self.browser.find_element(By.XPATH, '//*[@id="expanded-view-popup"]/div[4]/div[2]/div[2]/div/div[1]/div/div[2]/span[2]/span[2]').text 
            except NoSuchElementException:
                channel_name =''

            try:
                all_details = utils.get_check_table_df(self.browser.page_source, 'מספר בנק מחויב')
                comment = self.browser.find_element(By.XPATH, '//*[@id="expanded-view-popup"]/div[4]/div[2]/div[2]/div/div[2]/section/div[1]/div/div/div/div/span').text 
            except NoSuchElementException:
                all_details = utils.get_check_table_df(self.browser.page_source, 'מספר בנק מחויב')
            self.recent_transactions_data.append([datetime.now().strftime("%d/%m/%Y")] + data_row.text.split('\n') + [channel_name]+ all_details.iloc[:3][1].values.tolist() +["" ,""] + all_details.iloc[4:][1].values.tolist() + [comment])
            print("\t[img not found!]\n")
        except WebDriverException:
            print("\n\tCouldn't find the page [retrying]\n")
            self.get_no_image_table_data()

    def search_for_date_range_again(self):
        print("\n\tSearch and send request for_date_range_again...\n")
        WebDriverWait(self.browser,20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#advanced-search-window-btn > button"))).click()
        time.sleep(3)
        WebDriverWait(self.browser, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/main/section/div[4]/div[3]/div/div/div[2]/div/div[2]/div[1]/div/div[1]/advanced-search/div/div[2]/form/div[2]/button'))).click()
        time.sleep(3)

     
    def filter_data_for_recent_transactions(self):
        try:
            print("\n\n\t\t* * * * * * * * * * * * * Using Date Filter for Recent Transanctions * * * * * * * * * * * * *")
            time.sleep(3)
            WebDriverWait(self.browser,20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#advanced-search-window-btn > button"))).click()
            time.sleep(3)
            WebDriverWait(self.browser,20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="fromDate"]'))).click()
            try:
                from_month_year = self.browser.find_element(By.XPATH,'/html/body/div[2]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div[2]/div[1]/div/div[1]/advanced-search/div/div[2]/form/div[1]/div[1]/div/div[1]/div[1]/div[2]/div/div/ul/li/div/div/div/table/thead/tr[1]/th[2]').text.split()
            except WebDriverException:
                from_month_year = self.browser.find_element(By.XPATH,'/html/body/div[*]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div[2]/div[1]/div/div[1]/advanced-search/div/div[2]/form/div[1]/div[1]/div/div[1]/div[1]/div[2]/div/div/ul/li/div/div/div/table/thead/tr[1]/th[2]').text.split()

            diff_in_from_month = self.month_diff(from_month_year, self.start_date_transactions)
            self.select_transaction_date_from_calendar(diff_in_from_month, self.start_date_transactions)
            
            print('\tCALENDER MONTH YEAR:', from_month_year)
            print('\n\tDIFFERENCE IN MONTHS:', diff_in_from_month)
            time.sleep(3)
            # WebDriverWait(self.browser,20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="oshTransfersAdvancedSearchDateTO"]'))).click()
            # try:
            #     to_month_year = self.browser.find_element(By.XPATH,'/html/body/div[2]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div[2]/div[1]/div/div[1]/advanced-search/div/div[2]/form/div[1]/div[1]/div/div[1]/div[2]/div[2]/div/div/ul/li/div/div/div/table/thead/tr[1]/th[2]').text.split()
            # except WebDriverException:
            #     to_month_year = self.browser.find_element(By.XPATH,'/html/body/div[*]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div[2]/div[1]/div/div[1]/advanced-search/div/div[2]/form/div[1]/div[1]/div/div[1]/div[2]/div[2]/div/div/ul/li/div/div/div/table/thead/tr[1]/th[2]').text.split()
            # diff_in_to_month = self.month_diff(to_month_year, self.end_date_transactions)
            # self.select_transaction_date_from_calendar(diff_in_to_month, self.end_date_transactions)
            # time.sleep(4)
            
            try:
                WebDriverWait(self.browser, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div[2]/div[1]/div/div[1]/advanced-search/div/div[2]/form/div[2]/button'))).click()
            except WebDriverException:
                WebDriverWait(self.browser, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[*]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div[2]/div[1]/div/div[1]/advanced-search/div/div[2]/form/div[2]/button'))).click()
            time.sleep(5)
        except WebDriverException:
            print(WebDriverException, 'Error occurred!\nRetrying!!')
            self.browser.refresh()
            time.sleep(15)
            self.filter_data_for_recent_transactions()

    def filter_data_for_dep_checks(self):
        try:
            print("\n\n\t\t* * * * * * * * * * * * * Using Date Filter for Deposit Checks * * * * * * * * * * * * *")
            time.sleep(2)
            WebDriverWait(self.browser,20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#osh-checks-advanced-search-window-btn"))).click()
            time.sleep(2)
            WebDriverWait(self.browser,20).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='oshChecksAdvancedSearchFromDate']"))).click()
            # from_month_year = self.browser.find_element(By.XPATH,'/html/body/div[2]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div/div/div[1]/div[2]/section/section/div[1]/div[1]/div/div/div/ul/li/div/div/div/table/thead/tr[1]/th[2]').text.split()
            try:
                from_month_year = WebDriverWait(self.browser,10).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[2]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div/div/div[1]/div[2]/section/section/div[1]/div[1]/div/div/div/ul/li/div/div/div/table/thead/tr[1]/th[2]"))).text.split()
            except:
                from_month_year = WebDriverWait(self.browser,10).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[*]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div/div/div[1]/div[2]/section/section/div[1]/div[1]/div/div/div/ul/li/div/div/div/table/thead/tr[1]/th[2]"))).text.split()
            diff_in_from_month = self.month_diff(from_month_year, self.start_date_checks)
            self.select_check_date_from_calendar(diff_in_from_month, self.start_date_checks)
            print('\tCALENDER MONTH YEAR:', from_month_year)
            print('\n\tDIFFERENCE IN MONTHS:', diff_in_from_month)
            time.sleep(4)
            # WebDriverWait(self.browser,20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="oshChecksAdvancedSearchToDate"]'))).click()
            # # to_month_year = self.browser.find_element(By.XPATH,'//*[@id="main-content"]/div[2]/feature-flag/div/flag-off-component/div/div/div/div[1]/div[2]/section/section/div[1]/div[2]/div/div/div/ul/li/div/div/div/table/thead/tr[1]/th[2]').text.split()
            # try:
            #     to_month_year = WebDriverWait(self.browser,10).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[2]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div/div/div[1]/div[2]/section/section/div[1]/div[2]/div/div/div/ul/li/div/div/div/table/thead/tr[1]/th[2]"))).text.split()
            # except:
            #     to_month_year = WebDriverWait(self.browser,10).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[*]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div/div/div[1]/div[2]/section/section/div[1]/div[2]/div/div/div/ul/li/div/div/div/table/thead/tr[1]/th[2]"))).text.split()
            # diff_in_to_month = self.month_diff(to_month_year, self.end_date_checks)
            # self.select_check_date_from_calendar(diff_in_to_month, self.end_date_checks)
            # time.sleep(4)
            
            try:
                WebDriverWait(self.browser, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div/div/div[1]/div[2]/section/section/div[3]/button'))).click()
            except:
                WebDriverWait(self.browser, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[*]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div/div/div[1]/div[2]/section/section/div[3]/button'))).click()
        except WebDriverException:
            print(WebDriverException, 'Error Occured!\nRetrying!!')
            self.browser.refresh()
            time.sleep(15)
            self.filter_data_for_dep_checks()

    def select_check_date_from_calendar(self, diff, date):
        if date == self.start_date_checks:
            div = 1
            date_select = self.start_date_checks.split('/')[0]
            date_row = self.browser.find_elements(By.XPATH,"/html/body/div[*]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div/div/div[1]/div[2]/section/section/div[1]/div[1]/div/div/div/ul/li/div/div/div/table/tbody/tr")
        # if date == self.end_date_checks:
        #     div = 2
        #     date_select = self.end_date_checks.split('/')[0]
        #     date_row = self.browser.find_elements(By.XPATH,"/html/body/div[*]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div/div/div[1]/div[2]/section/section/div[1]/div[2]/div/div/div/ul/li/div/div/div/table/tbody/tr")

        for clicks in range(abs(diff)):
            time.sleep(4)
            if diff < 0:
                try:
                    WebDriverWait(self.browser,10).until(EC.element_to_be_clickable((By.XPATH, f"/html/body/div[2]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div/div/div[1]/div[2]/section/section/div[1]/div[{div}]/div/div/div/ul/li/div/div/div/table/thead/tr[1]/th[3]"))).click()
                except:
                    WebDriverWait(self.browser,10).until(EC.element_to_be_clickable((By.XPATH, f"/html/body/div[*]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div/div/div[1]/div[2]/section/section/div[1]/div[{div}]/div/div/div/ul/li/div/div/div/table/thead/tr[1]/th[3]"))).click()
            else:                                                                            
                try:
                    WebDriverWait(self.browser,10).until(EC.element_to_be_clickable((By.XPATH, f"/html/body/div[2]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div/div/div[1]/div[2]/section/section/div[1]/div[{div}]/div/div/div/ul/li/div/div/div/table/thead/tr[1]/th[1]"))).click()
                except:
                    WebDriverWait(self.browser,10).until(EC.element_to_be_clickable((By.XPATH, f"/html/body/div[*]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div/div/div[1]/div[2]/section/section/div[1]/div[{div}]/div/div/div/ul/li/div/div/div/table/thead/tr[1]/th[1]"))).click()
            
        try:
            for ridx,rows in enumerate(date_row):
                print(rows.text)
                if ridx == 0 and date_select > '07':
                    continue
                for cidx,td in enumerate(rows.text.split()):
                    if td == date_select:
                        time.sleep(4)
                        try:
                            WebDriverWait(self.browser,10).until(EC.element_to_be_clickable((By.XPATH, f"/html/body/div[2]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div/div/div[1]/div[2]/section/section/div[1]/div[1]/div/div/div/ul/li/div/div/div/table/tbody/tr[{ridx+1}]/td[{cidx+1}]/button"))).click()
                        except:
                            WebDriverWait(self.browser,10).until(EC.element_to_be_clickable((By.XPATH, f"/html/body/div[*]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div/div/div[1]/div[2]/section/section/div[1]/div[1]/div/div/div/ul/li/div/div/div/table/tbody/tr[{ridx+1}]/td[{cidx+1}]/button"))).click()
                        break
        
        except StaleElementReferenceException:
            pass

    def select_transaction_date_from_calendar(self, diff, date):
        if date == self.start_date_transactions:
            div = 1
            date_select = self.start_date_transactions.strftime('%d/%m/%Y')
            date_select = date_select.split('/')[0]
            date_row = WebDriverWait(self.browser, 20).until(EC.visibility_of_all_elements_located((By.XPATH,'/html/body/div[*]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div[2]/div[1]/div/div[1]/advanced-search/div/div[2]/form/div[1]/div[1]/div/div[1]/div[1]/div[2]/div/div/ul/li/div/div/div/table/tbody/tr')))

        # if date == self.end_date_transactions:
        #     div = 2
        #     date_select = self.end_date_transactions.split('/')[0]
        #     date_row = self.browser.find_elements(By.XPATH,'/html/body/div[*]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div[2]/div[1]/div/div[1]/advanced-search/div/div[2]/form/div[1]/div[1]/div/div[1]/div[2]/div[2]/div/div/ul/li/div/div/div/table/tbody/tr')
        for clicks in range(abs(diff)):
            time.sleep(4)
            if diff < 0:
                try:
                    WebDriverWait(self.browser,10).until(EC.element_to_be_clickable((By.XPATH, f"/html/body/div[2]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div[2]/div[1]/div/div[1]/advanced-search/div/div[2]/form/div[1]/div[1]/div/div[1]/div[{div}]/div[2]/div/div/ul/li/div/div/div/table/thead/tr[1]/th[3]"))).click()
                except WebDriverException:
                    WebDriverWait(self.browser,10).until(EC.element_to_be_clickable((By.XPATH, f"/html/body/div[*]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div[2]/div[1]/div/div[1]/advanced-search/div/div[2]/form/div[1]/div[1]/div/div[1]/div[{div}]/div[2]/div/div/ul/li/div/div/div/table/thead/tr[1]/th[3]"))).click()
            else:
                try:
                    WebDriverWait(self.browser,10).until(EC.element_to_be_clickable((By.XPATH, f"/html/body/div[2]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div[2]/div[1]/div/div[1]/advanced-search/div/div[2]/form/div[1]/div[1]/div/div[1]/div[{div}]/div[2]/div/div/ul/li/div/div/div/table/thead/tr[1]/th[1]"))).click()
                except:
                    WebDriverWait(self.browser,10).until(EC.element_to_be_clickable((By.XPATH, f"/html/body/div[*]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div[2]/div[1]/div/div[1]/advanced-search/div/div[2]/form/div[1]/div[1]/div/div[1]/div[{div}]/div[2]/div/div/ul/li/div/div/div/table/thead/tr[1]/th[1]"))).click()
                                                                                               
        try:
            for ridx,rows in enumerate(date_row):
                if ridx == 0 and date_select > '07':
                    continue
                for cidx,td in enumerate(rows.text.split()):
                    if td == date_select:
                        print(f'ridx:{ridx+1}, cidx:{cidx+1}')
                        time.sleep(4)
                        try:
                            WebDriverWait(self.browser,10).until(EC.element_to_be_clickable((By.XPATH, f"/html/body/div[2]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div[2]/div[1]/div/div[1]/advanced-search/div/div[2]/form/div[1]/div[1]/div/div[1]/div[{div}]/div[2]/div/div/ul/li/div/div/div/table/tbody/tr[{ridx+1}]/td[{cidx+1}]/button"))).click()
                        except:
                            WebDriverWait(self.browser,10).until(EC.element_to_be_clickable((By.XPATH, f"/html/body/div[*]/main/section/div[4]/div[3]/div/div/div[2]/feature-flag/div/flag-off-component/div/div[2]/div[1]/div/div[1]/advanced-search/div/div[2]/form/div[1]/div[1]/div/div[1]/div[{div}]/div[2]/div/div/ul/li/div/div/div/table/tbody/tr[{ridx+1}]/td[{cidx+1}]/button"))).click()
                        break
        
        except StaleElementReferenceException:
            pass

    def month_diff(self, month_year, sheet_date):
        current_date = datetime.now().strftime('%d/%m/%Y')
        sheet_date = sheet_date.strftime('%d/%m/%Y')
        curr_year = current_date.split('/')[2]
        month_dict = {'ינואר':1, 'פברואר':2, 'מרץ':3, 'אפריל':4, 'מאי':5, 'יוני':6, 'יולי':7, 'אוגוסט':8, 'ספטמבר':9, 'אוקטובר':10, 'נובמבר':11, 'דצמבר':12}
        if month_year[0] in month_dict:
            month_num = month_dict[month_year[0]]
        for month, month_idx in month_dict.items():
            if int(curr_year) != int(sheet_date.split('/')[2]):
                if month == month_year[0]:
                    year_diff = int(curr_year) - int(sheet_date.split('/')[2])
                    difference = (int(month_num) + (12*int(year_diff))) - int(sheet_date.split('/')[1])
            if int(curr_year) == int(sheet_date.split('/')[2]):
                if month == month_year[0]:
                    difference = int(month_num) - int(sheet_date.split('/')[1]) 
        return difference

    def upload_to_gdrive(self):
        print(' ---------------- UPLOAD IMAGE TO DRIVE ---------------- ')
        while True:
            access_token = utils.get_access_token()
            if utils.verify_token(access_token):
                break
            else:
                print('Invalid token, Retrying.')
        
        self.check_data_upload(access_token, self.gdrive_folder_id)
        self.transaction_data_upload(access_token, self.gdrive_folder_id)
        
    def transaction_data_upload(self, access_token, folder_id):
        for (idx, (ref, (front_img_path, back_img_path, pdf_name))), row in zip(enumerate(self.image_name.items()), self.recent_transactions_data):
            print(f'\n\nREF: {ref}, ROW: {row}\n')
            max_retry = 0
            row.insert(5, 'no image'); row.insert(6, 'no image'); row.insert(7, 'no pdf')
            time.sleep(2)
            if idx%20==0:
                self.browser.refresh()
            if front_img_path and back_img_path:
                print('\tImage and PDF Found!\n')
                print(pdf_name)
                while max_retry < 2:
                    try:
                        front_img_link = utils.upload_file_to_drive(front_img_path, access_token, folder_id)
                        back_img_link = utils.upload_file_to_drive(back_img_path, access_token, folder_id)
                        pdf_link = utils.upload_file_to_drive(pdf_name, access_token, folder_id)
                        print('front_img_link:', front_img_link)
                        print('back_img_link:', back_img_link)
                        print('pdf_link:', front_img_link)
                        break
                    except:
                        max_retry += 1
                        print('retry_n0:', max_retry)
                        print('Failed to get link for file, Retrying!')
                        time.sleep(5)
                        if max_retry == 2:
                            print('Setting [Failed] as default, unable to get link')
                            front_img_link = 'Failed'
                            back_img_link = 'Failed'
                            pdf_link = 'Failed'
                            break
                if row[4] == ref: 
                    row[5] = front_img_link; row[6] = back_img_link ; row[7] = pdf_link
                print('final row:', row)
            
        print('Recent Transactions Data\n', self.recent_transactions_data)

    def check_data_upload(self, access_token, folder_id):
        for idx, row in enumerate(self.dep_checks_data):
            max_retry = 0
            time.sleep(2)
            if row[9] and row[10]:
                while max_retry < 2:
                    try:
                        front_img_link = utils.upload_file_to_drive(row[9], access_token, folder_id)
                        back_img_link = utils.upload_file_to_drive(row[10], access_token, folder_id)
                        break
                    except:
                        max_retry += 1
                        print('Failed to get link for file, Retrying!')
                        time.sleep(5)
                        if max_retry == 2:
                            print('Setting [Failed] as default, unable to get link')
                            front_img_link = 'Failed'
                            back_img_link = 'Failed'
                            break
                row.pop(9); row.insert(9, front_img_link)
                row.pop(10); row.insert(10, back_img_link)
        print('Deposit Checks Data\n', self.dep_checks_data)

    def push_data_to_drive(self):
        print(f"\n\t[Pushing data to drive for user - {self.username}]")
        for data, sheet_name in self.all_data_map.items():
            print("\n\t\tPushing data for sheet:", sheet_name)
            task_data = utils.flattened_data(self, data)
            df = pd.DataFrame(task_data)
            if sheet_name == self.all_data_map['dep_checks_data']:
                df.to_csv(f'csv_{self.username}\{sheet_name}_{datetime.now().strftime("%d-%m-%Y_%H%M%S")}.csv')
            if sheet_name == self.all_data_map['recent_transactions_data']:
                df.to_csv(f'csv_{self.username}\{sheet_name}_{datetime.now().strftime("%d-%m-%Y_%H%M%S")}.csv')
            GoogleSheetHandler(data=task_data, sheet_name=sheet_name).appendsheet_records()

if __name__=='__main__':
    args = len(sys.argv)
    options = Options()

    if args > 1 and sys.argv[1].lower() == '--headless_mode=on':
        print('sys.argv:', sys.argv)
        """ Custom options for browser """ 
        prefs = {"download.default_directory" : "pdf/"}
        options.add_experimental_option("prefs",prefs)
        options.add_argument("--start-maximized")# maximized window
        options.headless = True
        browser = webdriver.Chrome(executable_path=config.DRIVER_PATH, options=options)
    else:
        ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
        prefs = {"download.default_directory" : os.path.join(ROOT_DIR, 'pdf\\', 'tmp\\')}
        options.add_experimental_option("prefs",prefs)
        options.add_argument("--start-maximized")# maximized window
        browser = webdriver.Chrome(executable_path=config.DRIVER_PATH, options=options)

    print("\n\t\t* *  * *  * *  * *  * *  * * START  * *  * *  * *  * *  * * *")
    action = ActionChains(browser)
    users = GoogleSheetHandler(sheet_name=config.SHEET_USERS_TAB_NAME).getsheet_records()
    
    for user in users[1:]:
        process_user = user[2]
        recent_transactions = user[3]
        dep_checks = user[4]
        letters = user[5]
        start_date_transactions = user[6]
        start_date_checks = user[7]
        start_date_letters = user[8]
        gdrive_folder_id = user[9]

        username, password = user[0], user[1]
        if process_user.upper() == 'NO':
            print(f'\n\t PROCESS USER {username}: {process_user.upper()}\n\n')
            print('\n\t SKIPPING USER\n\n')
            continue
        
        print(f'\n\t PROCESS USER {username}: {process_user.upper()}\n\n')
        print("\n\t START SCRAPPING FOR USER: %s\n\n" %username)

        last_transaction_date = datetime.strptime(start_date_transactions, '%d/%m/%Y')
        last_check_date = datetime.strptime(start_date_checks, '%d/%m/%Y')
        start_date_transactions = datetime.strptime(start_date_transactions, '%d/%m/%Y')
        current_date = datetime.now().date().strftime('%d/%m/%Y')

        print(f'\n\t GDRIVE FOLDER ID: {gdrive_folder_id}\n\n')

        check_date_diff = (datetime.strptime(current_date, '%d/%m/%Y') - last_check_date).days
        transaction_date_diff = (datetime.strptime(current_date, '%d/%m/%Y') - last_transaction_date).days

        flag = True
        check_flag = True
        while flag or check_flag:
            if last_transaction_date >= datetime.strptime(current_date, '%d/%m/%Y') or (flag==False and check_flag==False):
                print(f'\n\t ALL DATA ARE PROCESSED FOR USER: {username}\n\n')
                break
            print('*'*100)
            print(f'\n\t CURRENT DATE: {current_date}\t\n')
            print(f'\n\t START TRANSACTIONS DATE: {start_date_transactions}\t\n')
            if check_date_diff >= 360 and dep_checks.upper() == 'YES':
                print(f'\n\t YOU HAVE SELECTED CHECK DATE FOR 360+ DAYS!!\n\t SELECTING LAST 12 MONTHS CHECK DATA!\n')
            else: print(f'\n\t LAST CHECK DATE: {last_check_date}\t\n')
            if transaction_date_diff >= 360 and recent_transactions.upper() == 'YES':
                print(f'\n\t YOU HAVE SELECTED TRANSACTION DATE FOR 360+ DAYS!!\n\t SELECTING LAST 12 MONTHS TRANSACTION DATA!\n')
            else: print(f'\n\t LAST TRANSACTION DATE: {last_transaction_date}\t\n\n')            
            print('*'*100)
            
            scrapper = DataScrapping(
                browser, username, password, recent_transactions, dep_checks, letters, check_date_diff, transaction_date_diff,
                start_date_transactions, start_date_checks, start_date_letters, gdrive_folder_id
            )
            scrapper.login_to_site()
            if scrapper.user_login:
                try:
                    scrapper.create_directories()
                    check_flag = scrapper.get_dep_checks_data()
                    flag = scrapper.get_recent_transaction_data()
                except Exception as err:
                    print(f"Error in : {traceback.format_exc()}")      

                if check_flag == None: check_flag = True
                if flag == None: flag = False
                
                shutil.rmtree('pdf/tmp')
                shutil.rmtree(f'images_{username}/tmp')
                scrapper.upload_to_gdrive()
                scrapper.push_data_to_drive()
                scrapper.logout()

                print(f'\n\t All Checks Data [NOT] processed for {username}: {check_flag}\n\n')
                print(f'\n\t All Transaction Data [NOT] processed for {username}: {flag}\n\n')
            
            ref, transaction_date = utils.get_last_sheet_record(config.SHEET_DATA_TAB_NAME+'_'+username, username)
            if transaction_date:
                last_transaction_date = datetime.strptime(transaction_date, '%d/%m/%Y')
                start_date_transactions = last_transaction_date
                        
        print("\n\tEnd activity for user!\n\n")
    browser.close()

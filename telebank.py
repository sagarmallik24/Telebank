import os
import sys
import time
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
from selenium.common.exceptions import ElementNotInteractableException, StaleElementReferenceException, NoSuchElementException, WebDriverException, ElementClickInterceptedException

import utils
import config
from helpers.g_sheet_handler import GoogleSheetHandler

warnings.filterwarnings("ignore") 

class DataScrapping():
    
    def __init__(self, browser, username, password):
        self.browser = browser
        self.username = username
        self.password = password
        self.user_login = False

        self.dep_checks_data = []
        self.recent_transactions_data = []
        self.image_name = {}
        self.all_data_map = {'dep_checks_data': config.SHEET_CHECKS_TAB_NAME+'_'+self.username, 'recent_transactions_data': config.SHEET_DATA_TAB_NAME+'_'+self.username}

    def login_to_site(self):
        print(" Start user login..")
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
                print('Retrying... Login!')
                self.browser.delete_all_cookies()
                self.login_to_site()
            print('Login successfull')
        except:
            print("  Login Failed !!\n")
            time.sleep(3)

    def create_directories(self):
        print(f'\n\n\t\t* * * * * * * * * * Creating Directory for Images & PDF for USER: {self.username} * * * * * * * * * *')
        try:
            os.makedirs(f'images_{self.username}')
            print(f"\n\t\t'images_{self.username}' directory created for USER: {self.username}")
        except:
            print(f"\n\t\t'images_{self.username}' directory already exists for USER: {self.username}")
            pass
        try:
            os.makedirs('pdf')
            print("\n\t\tpdf directory created")
        except:
            print("\n\t\tpdf directory already exists")
            pass

    def logout(self):
        time.sleep(5)
        try:
            self.browser.find_element(By.CSS_SELECTOR, '#logOutLink').click()
            self.user_login = False
            time.sleep(5)
            print('Logged out user(%s) successfully!\n' %self.username)
            return self.user_login
        except ElementNotInteractableException or NoSuchElementException or StaleElementReferenceException or ElementClickInterceptedException:
            pass

    def get_dep_checks_data(self):
        date, ref = utils.get_last_sheet_record(self.all_data_map['dep_checks_data'], self.username)
        print(' Getting Deposit Checks Data.... ')
        time.sleep(3)
        self.browser.get('https://start.telebank.co.il/apollo/business/#/CHKVEW')
        self.browser.refresh()
        WebDriverWait(self.browser,10).until(EC.element_to_be_clickable((By.ID,"dropdownMenu2"))).click()
        WebDriverWait(self.browser,10).until(EC.element_to_be_clickable((By.LINK_TEXT, "12 חודשים אחרונים"))).click()
        time.sleep(9)
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
            if not df.empty:
                front_img_li, back_img_li = self.get_check_image(df.values.tolist())
                df['קישור לתמונה קדמית'] = front_img_li
                df['קישור לתמונה אחורה'] = back_img_li
            print(df)
            data = df.values.tolist()
            self.dep_checks_data = data
        except WebDriverException:
            print("\t\nCouldn't find the page [retrying]")
            self.get_dep_checks_data()

        except ValueError:
            print("\t\nNo table found between selected dates!")
            return

    def get_check_image(self, data):
        data_li = [row[5:7] for row in data]
        front_img_li = []
        back_img_li = []
        row = 0
        try:
            for data_row in self.browser.find_elements(By.XPATH, '//*[@id="main-content"]/div[2]/div/div/div/div[2]/div[1]/div[2]/div/div/section/table/tbody/tr'):
                print('Processing Data for Check No: ', data_row.text.split(' ')[10])
                if int(data_row.text.split(' ')[10]) == data_li[row][0]:
                    time.sleep(3)
                    data_row.click()
                    check_number = data_row.text.split(' ')[10]
                    time.sleep(3)
                    front_image = self.browser.find_element(By.CSS_SELECTOR, 'body > div.modal.discountBiz-modal-general.cs-spa-sme-content.topbar-modal.checksExpand.fade.ng-scope.ng-isolate-scope.in > div > div > div > div.containerOsh.container-fluid > div.contentOsh.row > section.col-sm-6.col-print-9 > div > div > img').get_attribute('src')
                    WebDriverWait(self.browser,10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'body > div.modal.discountBiz-modal-general.cs-spa-sme-content.topbar-modal.checksExpand.fade.ng-scope.ng-isolate-scope.in > div > div > div > div.containerOsh.container-fluid > div.contentOsh.row > section.col-sm-6.col-print-9 > button > img'))).click()
                    time.sleep(2)
                    back_image = self.browser.find_element(By.CSS_SELECTOR, 'body > div.modal.discountBiz-modal-general.cs-spa-sme-content.topbar-modal.checksExpand.fade.ng-scope.ng-isolate-scope.in > div > div > div > div.containerOsh.container-fluid > div.contentOsh.row > section.col-sm-6.col-print-9 > div > div > img').get_attribute('src')
                    front_img_name = f'images_{self.username}/{check_number.replace("/", "-")}_1.jpg'
                    urllib.request.urlretrieve(front_image, front_img_name)
                    front_img_li.append(front_img_name)
                    back_img_name = f'images_{self.username}/{check_number.replace("/", "-")}_2.jpg'
                    urllib.request.urlretrieve(back_image, back_img_name)
                    back_img_li.append(back_img_name)
                    time.sleep(2)
                    self.browser.find_element(By.CSS_SELECTOR, 'button[type="button"]').click()
                    row += 1
                else:
                    print('Data Already Exists!')
            return front_img_li, back_img_li
        except WebDriverException:
            print("Couldn't find the page [retrying]")
            self.get_check_image()
   
    def get_reference_index(self, ref):
        print('Getting last Saved Index')
        count = 0
        while True:
            refs =  self.browser.find_element(By.XPATH, f'//*[@id="lastTransactionTable-cell-{count}-3"]')
            count += 1 
            if refs.text == ref:
                break
        return count

    def get_recent_transaction_data(self):
        ref, date = utils.get_last_sheet_record(self.all_data_map['recent_transactions_data'], self.username)
        print(' Getting Image Data from Last Transactions ')
        time.sleep(2)
        self.browser.get('https://start.telebank.co.il/apollo/business/#/OSH_LENTRIES_ALTAMIRA')
        self.browser.refresh()
        time.sleep(9)
        if ref:
            last_row_idx = int(self.get_reference_index(ref[-1]))
        if not ref:
            ref.append(0)
        time.sleep(3)
        print('Please Wait! Skipping to the Last Saved Index......')

        try:
            for idx, data_row in enumerate(self.browser.find_elements(By.CSS_SELECTOR, '.rc-table-row-content')):
                if idx < last_row_idx:
                    continue

                row_date, row_ref, check_no = data_row.text.split('\n')[0], data_row.text.split('\n')[3], data_row.text.split()[3].split(':')[1]
                print("\tprocess row data: ref:", data_row.text.split('\n')[3])
                if row_ref.split('/')[0] not in ref:
                    WebDriverWait(self.browser,20).until(EC.element_to_be_clickable((By.ID,f"lastTransactionTable-row-{idx}"))).click()
                    time.sleep(9)
                    self.unique_reference = WebDriverWait(self.browser,10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, '.text-box-value-color.ng-binding.number-font'))).text
                    img_element = utils.verify_element(browser=self.browser, by_selector=By.XPATH, path='.//*[@id="single-check-view-con"]/div/div[1]/img[1]')
                    self.image_name[self.unique_reference] = ['', '']
                    if img_element:
                        print('img found!')
                        time.sleep(1)
                        ''' Downloading PDF '''
                        self.browser.find_element(By.XPATH, '//*[@id="expanded-view-popup"]/div[4]/div[1]/div/div/div/ul/li[2]/a').click()
                        time.sleep(17)  
                        self.browser.implicitly_wait(17)
                        pdfs = os.listdir('pdf')
                        pdf_ref = row_ref.split('/')[0]
                        new_pdf_name = f'{pdf_ref}_{self.username}_{check_no}.pdf'
                        utils.rename_file('pdf', pdfs[int(idx-last_row_idx)], new_pdf_name) #Renaming pdf file 
                       
                        self.download_image(check_no)
                        self.get_image_table_data(data_row, check_no)
                    else:
                        self.get_no_image_table_data(data_row)
                    try:
                        self.browser.find_element(By.CSS_SELECTOR, 'button[type="button"]').click();time.sleep(1.5)
                    except ElementNotInteractableException or StaleElementReferenceException:
                        print('Not Interactable Image: [IGNORE]')
                        pass
                else:
                    print(f'Data Already Exists')
                    continue
        except WebDriverException as err:
            print(f"\t\n Error Occured Due to: {err} ")
            print("\t\n ** ** ** **  Uploading the extracted imgs & pdf to gdrive and Pushing data to sheet till now ** ** ** **")
  
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

    def get_image_table_data(self, data_row):
        channel_name = self.browser.find_element(By.XPATH, '//*[@id="expanded-view-popup"]/div[4]/div[2]/div[2]/div/div[1]/div/ng-include/div/div[2]/span[2]/span[2]').text
        all_details = utils.get_check_table_df(self.browser.page_source, 'מספר בנק')
        front_img_name, back_img_name = self.download_image()
        self.image_name[self.unique_reference] = [front_img_name, back_img_name]
        self.recent_transactions_data.append([datetime.now().strftime("%d/%m/%Y")] + data_row.text.split('\n') + [channel_name]+ all_details.iloc[:][1].values.tolist() +["" , "", ""])

    def get_no_image_table_data(self, data_row):
        try:
            comment = ''
            try:
                channel_name = self.browser.find_element(By.XPATH, '//*[@id="expanded-view-popup"]/div[4]/div[2]/div[2]/div/div[1]/div/div[2]/span[2]/span[2]').text
            except NoSuchElementException:
                channel_name = '' 
            try:
                all_details = utils.get_check_table_df(self.browser.page_source, 'מספר בנק מחויב')
                comment = self.browser.find_element(By.XPATH, '//*[@id="expanded-view-popup"]/div[4]/div[2]/div[2]/div/div[2]/section/div[1]/div/div/div/div/span').text 
            except NoSuchElementException:
                all_details = utils.get_check_table_df(self.browser.page_source, 'מספר בנק מחויב')
            self.recent_transactions_data.append([datetime.now().strftime("%d/%m/%Y")] + data_row.text.split('\n') + [channel_name]+ all_details.iloc[:3][1].values.tolist() +["" ,""] + all_details.iloc[4:][1].values.tolist() + [comment])
            print("\t[img not found!]\n")
        except WebDriverException:
            print("couldn't find the page [retrying]")
            self.get_no_image_table_data()

    def upload_to_gdrive(self):
        print(' ---------------- UPLOAD IMAGE TO DRIVE ---------------- ')
        while True:
            access_token = utils.get_access_token()
            if utils.verify_token(access_token):
                break
            else:
                print('Invalid token, Retrying.')
        
        pdfs = os.listdir('pdf')
        sorted(filter(os.path.isfile, pdfs), key=os.path.getmtime)
        for idx, row in enumerate(self.dep_checks_data):
            if row[9] and row[10]:
                front_img_link = utils.upload_file_to_drive(row[9], access_token)
                back_img_link = utils.upload_file_to_drive(row[10], access_token)
                row.pop(9); row.insert(9, front_img_link)
                row.pop(10); row.insert(10, back_img_link)
        print(self.dep_checks_data)

        for (idx, (ref, (front_img_path, back_img_path))), row in zip(enumerate(self.image_name.items()), self.recent_transactions_data):
            row.insert(5, 'no image'); row.insert(6, 'no image'); row.insert(7, 'no pdf')
            if front_img_path and back_img_path:
                print('image and pdf found!')
                front_img_link = utils.upload_file_to_drive(front_img_path, access_token)
                back_img_link = utils.upload_file_to_drive(back_img_path, access_token)
                try:
                    pdf_link = utils.upload_file_to_drive(pdfs[idx], access_token)
                except IndexError:
                    reference = front_img_path.split('/')[1].split('_')[0]
                    print(f'\n\tNo Pdf Found for the reference No. {reference}')
                    continue                
                print('row', row[4])
                if row[4] == ref:
                    row[5] = front_img_link; row[6] = back_img_link ; row[7] = pdf_link
        print('self.recent_transactions_data\n', self.recent_transactions_data)

    def push_data_to_drive(self):
        print(f"\n\t[Pushing data to drive for user - {self.username}]")
        for data, sheet_name in self.all_data_map.items():
            print("\n\t\tPushing data for sheet:", sheet_name)
            task_data = utils.flattened_data(self, data)
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
        prefs = {"download.default_directory" : os.path.join(ROOT_DIR, 'pdf\\')}
        options.add_experimental_option("prefs",prefs)
        options.add_argument("--start-maximized")# maximized window

        browser = webdriver.Chrome(executable_path=config.DRIVER_PATH, options=options)

    print(" * *  * *  * *  * *  * *  * * START  * *  * *  * *  * *  * * ")
    action = ActionChains(browser)
    users = GoogleSheetHandler(sheet_name=config.SHEET_USERS_TAB_NAME).getsheet_records()
    for user in users[1:]:
        username, password = user[0], user[1]
        print("Start scrapping for user: %s" %username)
        scrapper = DataScrapping(browser, username, password)
        scrapper.login_to_site()

        if scrapper.user_login:
            try:
                scrapper.create_directories()

                scrapper.get_dep_checks_data()
                scrapper.get_recent_transaction_data()
            except Exception as err:
                print(f"Error in : {traceback.format_exc()}")           
           
            scrapper.upload_to_gdrive()
            scrapper.push_data_to_drive()
            scrapper.logout()
            
        print("End activity for user!\n\n")
    browser.close()
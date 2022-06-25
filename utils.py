import os
import json
import config
import shutil
import requests
import httplib2
import pandas as pd
from datetime import datetime
from bs4 import BeautifulSoup
from oauth2client import client, GOOGLE_TOKEN_URI
from helpers.g_sheet_handler import GoogleSheetHandler
from selenium.common.exceptions import NoSuchElementException

def get_table_df(page_source, table_id):
    soup = BeautifulSoup(page_source, 'html.parser')
    tables = soup.find('table', attrs={"class":table_id})
    df = pd.read_html(str(tables))[0].dropna(how='all')
    return df.fillna('')

def get_check_table_df(page_source, table_id):
    soup = BeautifulSoup(page_source, 'html.parser')
    try:
        tables = soup.find(text = table_id).find_parent('table')
        df = pd.read_html(str(tables))[0].dropna(how='all')
        return df.fillna('')
    except AttributeError:
        return pd.DataFrame(columns=range(5))

def verify_token(access_token):
    headers = {"Authorization": f"Bearer {access_token}"}
    r = requests.post("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart", headers=headers)
    if r.json().get('error'):
        return False
    return True

def flattened_data(scrapper, data):
    if data == 'dep_checks_data':
        result = scrapper.dep_checks_data
    if data == 'recent_transactions_data':
        result = scrapper.recent_transactions_data
    return result

def verify_element(browser, by_selector, path):
    try:
        browser.find_element(by_selector, path)
    except NoSuchElementException:
        return False
    return True


def upload_file_to_drive(file, access_token, GDRIVE_IMAGE_FOLDER_ID):
    print(f'\tfile:{file}')
    headers = {"Authorization": f"Bearer {access_token}"}
    if 'jpg' in file:
        para = { "name": f"{file.replace('images/', '')}", "parents": [GDRIVE_IMAGE_FOLDER_ID]}
        files = {
            'data': ('metadata', json.dumps(para), 'application/json; charset=UTF-8'),
            'file': open(f'{file}', "rb")
        }
    if 'pdf' in file:
        para = { "name": f"{file}", "parents": [GDRIVE_IMAGE_FOLDER_ID]}
        files = {
            'data': ('metadata', json.dumps(para), 'application/json; charset=UTF-8'),
            'file': open(f'pdf/{file}', "rb")
        }

    r = requests.patch(
        f"https://www.googleapis.com/upload/drive/v3/files/{files}", #?uploadType=multipart
        headers = headers #, files = files
    )

    if r.text == 'Not Found':
        r = requests.post(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        headers = headers, files = files, timeout = 10
    )
    # print(r.json())
    print("\t", r.json()['name'], r.json()['id'])
    file_link = 'https://drive.google.com/file/d/' + r.json()['id']
    return file_link


def get_last_sheet_record(sheet_name, username):
    print(f"\n\t[Getting last sheet record]")
    if sheet_name == config.SHEET_CHECKS_TAB_NAME+'_'+username:
        check_li = GoogleSheetHandler(sheet_name=sheet_name).getsheet_records()
        try:
            date_li = [datetime.strptime(dates[1], '%d/%m/%Y').date() for dates in check_li[1:] if any(x for x in dates)]
            print('date_li:', date_li)
            ref_li = [refs[3].split('/')[0] for refs in check_li[1:] if any(x for x in refs)]
            print('ref_li:', ref_li)
            date, ref = max(date_li).strftime('%d/%m/%Y'), max(ref_li)
        except:
            date, ref = '', ''
        return date, ref
    if sheet_name == config.SHEET_DATA_TAB_NAME+'_'+username:
        txn_li = GoogleSheetHandler(sheet_name=sheet_name).getsheet_records()
        try:
            date_li = [datetime.strptime(dates[1], '%d/%m/%Y').date() for dates in txn_li[1:] if any(x for x in dates)]
            ref_li = [refs[4].split('/')[0] for refs in txn_li[1:] if any(x for x in refs)]
            while '' in date_li: date_li.remove('')
            while '' in ref_li: ref_li.remove('')
            print('date_li:', date_li)
            print('ref:', ref_li)
            date, ref = max(date_li).strftime('%d/%m/%Y'), ref_li
        except:
            print("t\n No data in sheet!")
            date, ref = '', ''
        return ref, date

def get_access_token():
    file = open(config.CLIENT_CRED_FILE)
    data = json.load(file)
    credentials = client.OAuth2Credentials(
        access_token = None, 
        user_agent = "user-agent: google-oauth-playground",
        client_id = data['web']['client_id'],
        client_secret = data['web']['client_secret'], 
        refresh_token = config.REFRESH_TOKEN, 
        token_expiry = None, 
        token_uri = GOOGLE_TOKEN_URI,
        revoke_uri= None
        )

    credentials.refresh(httplib2.Http())
    access_token = credentials.access_token
    return access_token
    
def parse_date(date):
    date = date.split("/")
    return date[2]+date[1]+date[0]
        
def rename_file(path, filename, file_new_name):
    os.chdir(path)
    os.rename(filename, file_new_name)
    os.chdir('../../')

def create_dir(dir):
    print(f"\n\tRemoving and Creating DIR: {dir}\n")
    if os.path.isdir(dir):
        shutil.rmtree(dir)        
    os.makedirs(dir+'/tmp')

def get_check_no(ref_str):
    check_no = 'dummy'
    if (":" in  ref_str) and ref_str.split(":")[1]:
        check_no = ref_str.split(":")[1]
    return check_no


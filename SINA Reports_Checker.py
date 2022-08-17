import re
from datetime import datetime
import  http
from http import client
from random import randrange
import time
import bs4
from pymongo import UpdateOne, MongoClient
from textblob import TextBlob
from urllib.parse import urlparse
from bs4 import BeautifulSoup
import requests

connect = MongoClient('IP_address', port_number)
secMongodb = db_name
type_extracted_rows = []

CHSE_base_url = 'https://vip.stock.finance.sina.com.cn/corp/view/vCB_BulletinGather.php'
chse_types = {
                'ndbg':{'TypeName':'Annual Report','TypeKey':'AR'},
                'ndbgzy':{'TypeName':'Annual Report (Summary)','TypeKey':'ARS'},
                'sjdbg':{'TypeName':'Third Quarter Report','TypeKey':'TQR'},
                'sjdbgzy':{'TypeName':'Third Quarter Report (Summary)','TypeKey':'TQRS'},
                'yjdbg':{'TypeName':'First Quarter Report','TypeKey':'FQR'},
                'yjdbgzy':{'TypeName':'First Quarter Report (Summary)','TypeKey':'FQRS'},
                'zqbg': {'TypeName': 'Second Quarter Report', 'TypeKey': 'SQR'}
              }

headers = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Accept-Language": "en-US,en;q=0.9,ar;q=0.8",
    "Cache-Control": "max-age=0",
    "Connection": "keep-alive",
    "Host": "vip.stock.finance.sina.com.cn",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-User": "?1",
    "Upgrade-Insecure-Requests": "1",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36",
}
#http_proxy = "http://192.168.7.182:808"
#https_proxy = "https://192.168.7.182:808"


def insert_CHSE_data_to_mongo(data):
    try:
        collection = secMongodb.collection_name
        operations = []
        for doc in data:
            operations.append(UpdateOne({"Original HTML Link": doc["Original HTML Link"]}, { "$set": doc } , upsert=True))
        result = collection.bulk_write(operations , ordered=False)
    except Exception as e:
        print("Oops!", e, "occurred.")

def Generate_Full_URL(Source_URL , url):
    scheme = urlparse(Source_URL).scheme
    Domain = urlparse(Source_URL).netloc
    Full_URL = url
    if url.startswith('http'):
        Full_URL = url
    elif url[:2] == '//':
        Full_URL = scheme + ':' + url
    elif url.startswith('../'):
        Full_URL = scheme + '://' + Domain + url.replace('..', '')
    elif url.startswith('..'):
        Full_URL = scheme + ':' + url.replace('..', '//')
    elif url.startswith('./'):
        Full_URL = scheme + ':' + url.replace('./', '//')
    elif url.startswith('/'):
        Full_URL = scheme + '://' + Domain + url
    #elif url and url.lstrip()[0].isalpha():
     #   Full_URL = scheme + '://' + Domain + url
    elif url.startswith('/ - /'):
        Full_URL = scheme + '://' + Domain + url.replace('/ - /', '/')
    else:
        Full_URL = scheme + '://' + Domain + '/' + url
        #print("Stop")
    return Full_URL

def Chineese_Checker():
    print("-------------------------------------- Start Get CHSE Type Feeds ---------------------------------------------")
    # Loop over types
    for type in list(chse_types.keys()):
        print("--- Type --- ", type)
        for page_indx in range(1, 50):
            # Generate Type Url
            base_url_type = CHSE_base_url +'?ftype=' + str(type)
            curr_url_type = base_url_type+'&gg_date=2020-10-28' + '&page_index=' + str(page_indx)
            #curr_url_type = 'https://vip.stock.finance.sina.com.cn/corp/view/vCB_BulletinGather.php?stock_str=%B4%FA%C2%EB%2F%C3%FB%B3%C6%2F%C6%B4%D2%F4&gg_date=2020-02-17&ftype=ndbg'
            try:
                print(curr_url_type)
                htmlinner = requests.get(curr_url_type, allow_redirects=True, headers=headers).content
                soupinner = BeautifulSoup(htmlinner, 'html.parser')
                rows = soupinner.find("table", {"class": "body_table"}).find('tbody').find_all('tr')
                no_records = False
                # Loop over rows
                for r in rows:
                    if all( kw in str(TextBlob(r.text).translate(to='en')).lower() for kw in ['no','records','sorry']):
                        no_records = True
                        break
                    # Get all a tags
                    r_aTags = r.find_all('a')
                    # report url
                    r_lnk = r_aTags[0].attrs['href']
                    seq_no = r_lnk.split('id=')[-1]
                    # Get Report Info
                    print("----*", seq_no,"----*")
                    rep = {}
                    #rep['sequentialno'] = seq_no
                    rep['ReportName'] = str(TextBlob(r_aTags[0].text).translate(to='en'))
                    print("---",rep['ReportName'])
                    rep['ReportType'] = chse_types[type]['TypeName']
                    rep['ReportTypeKey'] = chse_types[type]['TypeKey']
                    # Get date td
                    rep['Date'] = datetime(2020, 10, 28, 0, 0, 0)
                    """r_tds = list(
                        filter(lambda x: re.match(r'\d{4}-\d{1,2}-\d{1,2}', x.text) is not None, r.find_all('td')))
                    if len(r_tds) > 0:
                        rep['Date'] = datetime.datetime.strptime(r_tds[-1].text, '%Y-%m-%d')"""
                    rep['Original HTML Link'] = Generate_Full_URL(CHSE_base_url, r_lnk)
                    rep['Original PDF Link'] = ''
                    pdf_lnks = list(filter(lambda a: all(word in str(a).lower() for word in ['href', '.pdf']), r_aTags))
                    if len(pdf_lnks) > 0:
                        rep['Original PDF Link'] = Generate_Full_URL(CHSE_base_url, pdf_lnks[0].attrs['href'])
                    rep['Translated HTML Link'] = 'https://translate.google.com/translate?depth=1&hl=en&nv=1&prev=search&rurl=translate.google.com&sl=zh-CN&sp=nmt4&u=' + rep['Original HTML Link']
                    rep['Translated PDF Link'] =  'https://translate.google.com/translate?depth=1&hl=en&nv=1&prev=search&rurl=translate.google.com&sl=zh-CN&sp=nmt4&u=' + rep['Original PDF Link']
                    rep['PDF Link Status'] = '1'
                    rep['Status'] = '1'
                    rep['Tables'] = '0'
                    rep['Exception'] = ''
                    type_extracted_rows.append(rep)
                # sleep time
                time.sleep(randrange(12))
                if no_records:
                    break
            except Exception as e:
                print("OOPS!! General Error")
                print(str(e))
                time.sleep(randrange(10))
    insert_CHSE_data_to_mongo(type_extracted_rows)

Chineese_Checker()

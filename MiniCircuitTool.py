import urllib
import os.path
from datetime import datetime
import pandas as pd
from bs4 import BeautifulSoup
import warnings
import time
import ctypes
from urllib.parse import urlparse
from selenium import webdriver
from selenium.webdriver.common.by import By
#______________________________________________________#
#----- Chrome Driver-----#
driver = None
proxy = None
chrome_options = webdriver.ChromeOptions()
chrome_path = r'chromedriver.exe'
#___________________ Classes and Functions ___________________#
###################################################################
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
#-------------------- Insert to file---------#
def insert_to_file(data_list,lnks_status_file):
    try:
        with open(lnks_status_file, "a",encoding='utf-8') as resultFile:
            for data in data_list:
                #print(data)
                data = str(data).replace('\n','').replace('\t','').strip()
                resultFile.write(data + '\t')
            #print("---------")
            #print(data_list)
            #print("---------")
            resultFile.write('\n')
    except Exception as ex:
        print("insert_to_file() >>  Exception: ",ex)
#------------------- add part data to file-----------------------------#
def data_tofile(cir_lnk, child_lnks,lnk_status,status,html_file_path, lnks_status_file):
    try:
        if len(child_lnks) > 0:
            for lnk in child_lnks:
                child_data = []
                child_lnk, child_description = lnk
                child_data.extend((cir_lnk, child_lnk, child_description, lnk_status, status,html_file_path))
                # print(Fore.CYAN, part_data_list, Fore.RESET)
                insert_to_file(child_data, lnks_status_file)
        else:
            child_data = []
            child_data.extend((cir_lnk, '-', '-', lnk_status, status,html_file_path))
            # ---------------- Insert to file-------------------#
            insert_to_file(child_data, lnks_status_file)
    except Exception as ex:
        print('data_tofile >> Exception: ',ex)
#--------------#
def get_tags(soup,cir_lnk,lnks_status_file):
    lnk_childs_lst = []
    alert_item = part_data = None
    try:
        alert_item = soup.find(id='myModal')
        lnk_alert_childs = alert_item.find_all('a')
        lnk_childs_lst.extend([(ch.text, Generate_Full_URL('https://www.minicircuits.com/', ch.get('href'))) for ch in lnk_alert_childs if ch.get('href') is not None])
    except:
        print("get_tags() >> No Alert Item")
    try:
        data_area = soup.find(id='content_area_home')
        lnk_childs = data_area.find_all('a')
        lnk_childs_lst.extend([(ch.text, Generate_Full_URL('https://www.minicircuits.com/', ch.get('href'))) for ch in lnk_childs if ch.get('href') is not None])
    except Exception as ex:
        print("get_tags() >> Exception: ",ex)
    finally:
        return lnk_childs_lst
#-------------------------- Save Html Page --------------#
def save_html(soup,cir_lnk):
    html_saved = False
    html_file_path = ''
    try:
        html_name = cir_lnk.split('?model=')[-1] + '.html'
        html_file_path = os.path.join(html_files_path, html_name)
        with open(html_file_path, "w",encoding='utf-8') as fp:
             fp.write(str(soup))
        html_saved = True
    except Exception as ex:
        print("save_html() >> Link: ",cir_lnk)
    finally:
        return (html_saved,html_file_path)
#-------------------- extract data frm lnks-----------------#
def start_extractlinks(links_lst,lnks_status_file):
    for cir_lnk in links_lst:
        part_data = None
        lnk_status = status =  html_saved =  html_file_path = ''
        print("_____",cir_lnk,"______")
        try:
            ######################## new
            try:
                driver.get(cir_lnk)
            except BaseException as ex:
                print(">>> Request Url: ", str(ex))
            #############
            try:
                part_data = driver.find_element(by=By.ID, value='wrapper')
            except: pass
            if part_data is not None:
                print("### Success")
                lnk_status = '3'
                status = 'successful'
                soup = BeautifulSoup(driver.page_source, features="html.parser")
                # -------------- Save Html Page --------------------------------------#
                html_saved, html_file_path = save_html(soup,cir_lnk)
                #--------------- Find <a> tags -------------------------------------#
                child_lnks= get_tags(soup, cir_lnk, lnks_status_file)
                #------------- Save child_lnks into Excel -------------------#
                data_tofile(cir_lnk, child_lnks,lnk_status,status,html_file_path, lnks_status_file)
            else:
                lnk_status = '4'
                status = 'blocked'
                print("### Blocked")
                data_tofile(cir_lnk, '-', lnk_status, status,html_file_path, lnks_status_file)
            time.sleep(30)
        except Exception as ex:
            lnk_status = '4'
            status = 'blocked'
            data_tofile(cir_lnk, '-', lnk_status, status,html_file_path, lnks_status_file)
            print("Get Link Page >> Exception: ", ex)
            time.sleep(30)
################# Program Start #################
#----- Send Start Message -------#
ctypes.windll.user32.MessageBoxW(None, "Please wait until the current process is completed", "Start Message",0x40, 0)
warnings.filterwarnings('ignore')
#___________________Definitions__________________#
#-------Link Error list --------
lnks_err_lst = []
#--------Excel links List-----------#
links_lst = []
# ----- Get Tbl Header------#
main_header_lst = ["MiniCir Link",'Child Link Title','Child Link','Status','Exception','MiniCir Link Html Path']
headers = {"Pragma": "no-cache", 'Cache-Control': 'no-cache', 'referer': 'minicircuits.com',
               'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36'}
#################################### Start Execution###################################
#---------------- Get Links from last modified sheet ------------------------#
miniCi_tool_dir = os.path.dirname(os.getcwd())
#print(path_getExcelSheet_wthTool_dir)
all_subdirs = [os.path.join(miniCi_tool_dir, d) for d in os.listdir(miniCi_tool_dir) if os.path.isdir(miniCi_tool_dir)]
files_xlsx = [f for f in all_subdirs if f.lower().endswith('xlsx') ]
latest_file = max(files_xlsx, key=os.path.getmtime) if len(files_xlsx)>0 else 0
lnks_status_file = ''
#------ for html_files dir ------#
now = datetime.now()
date_time = now.strftime("%m-%d-%Y_%H-%M-%S")
htmllocal_dir = "Links-Html_" + str(date_time)
html_files_path = os.path.join(miniCi_tool_dir, htmllocal_dir)
#------------------- Read Excel Sheet ------------------#
if latest_file==0:
    ctypes.windll.user32.MessageBoxW(None, "Excel Sheet Not Found", "Error", 0x10, 0)
else:
    print("Excel File >>>> ",latest_file)
#####################################################################
if __name__ == '__main__':
    #--------------- Start ----------------------#
    if latest_file !=0 :
        df_reqLinks = pd.read_excel(latest_file)
        links_lst = list(df_reqLinks.iloc[:, 0])
        # ---------------  Start Extract Child_Lnks from links -----------------------#
        if len(links_lst) > 0:
            # ---------------------- Create Result Folder and TextFile--------------------#
            os.mkdir(html_files_path)
            lnks_status_file = os.path.join(html_files_path , "Minicircuits.txt")
            with open(lnks_status_file, "a",encoding='utf-8') as resultFile:
                for header in main_header_lst:
                    resultFile.write(header + '\t')
                resultFile.write('\n')
            prefs = {"download.default_directory": html_files_path,  # IMPORTANT - ENDING SLASH V IMPORTANT
                     "directory_upgrade": True}
            chrome_options.add_experimental_option("prefs", prefs)
            driver = webdriver.Chrome(executable_path=chrome_path, chrome_options=chrome_options)
            start_extractlinks(links_lst, lnks_status_file)
            driver.quit()
        else:
            ctypes.windll.user32.MessageBoxW(None, "No links in Excel sheet", "Error", 0x10, 0)
    # ----------- Close Log File ---------------------#
    try:
        lnks_status_file.close()
    except:
        pass
    if latest_file != 0:
       ctypes.windll.user32.MessageBoxW(None, "Please Check folder \n\n" + str(html_files_path) + "", "Completed", 0x40,0)
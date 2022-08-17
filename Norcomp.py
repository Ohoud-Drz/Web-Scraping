##################################### Config ######################################
# Proxies
PROXY_lst = []
VALID_PROXY_LST = []
##################################### Main #########################################
#__________________________________ Libraries _______________________________________#
import pandas as pd
from datetime import datetime, timedelta
import time
from selenium import webdriver
from colorama import Fore
from selenium.webdriver.common.by import By
import ctypes
import glob
import os
import warnings
from urllib.parse import urlparse, urljoin
from selenium.webdriver.support.ui import Select
import urllib.request
################################# Global Variabes ################################################
#----- Chrome Driver-----#
driver = None
proxy = None
chrome_options = webdriver.ChromeOptions()
chrome_path = r'chromedriver.exe'
#################################### Start Execution###################################
#----- Send Start Message -------#
ctypes.windll.user32.MessageBoxW(None, "Please wait until the current process is completed", "Start Message",0x40, 0)
warnings.filterwarnings('ignore')
#------------------------------------------------#
main_header_lst = ["Part Number", 'Status','Exception','Compilance Doc','Pdf Path']
#---------------- Get Links from last modified sheet ------------------------#
miniCi_tool_dir = os.path.dirname(os.getcwd())
#print(path_getExcelSheet_wthTool_dir)
all_subdirs = [os.path.join(miniCi_tool_dir, d) for d in os.listdir(miniCi_tool_dir) if os.path.isdir(miniCi_tool_dir)]
files_xlsx = [f for f in all_subdirs if f.lower().endswith('xlsx') ]
latest_file = max(files_xlsx, key=os.path.getmtime) if len(files_xlsx)>0 else 0
parts_status_file = ''
#------------------- Read Excel Sheet ------------------#
if latest_file==0:
    ctypes.windll.user32.MessageBoxW(None, "Excel Sheet Not Found", "Error", 0x10, 0)
else:
    print("Excel File >>>> ",latest_file)
    df_reqLinks = pd.read_excel(latest_file)
    parts_lst = list(df_reqLinks.iloc[:, 0])
    # --------------------------------------#
    if len(parts_lst) > 0:
        # ---------------------- Create Result Folder and TextFile--------------------#
        now = datetime.now()
        date_time = now.strftime("%m-%d-%Y_%H-%M-%S")
        pdflocal_dir = "pdfs_" + str(date_time)
        pdf_files_path = os.path.join(miniCi_tool_dir, pdflocal_dir)
        os.mkdir(pdf_files_path)
        parts_status_file = pdf_files_path + "/Norcomp_LogFile.txt"
        with open(parts_status_file, "a",encoding='utf-8') as resultFile:
            for header in main_header_lst:
                resultFile.write(header + '\t')
            resultFile.write('\n')
    else:
        ctypes.windll.user32.MessageBoxW(None, "No links in Excel sheet", "Error", 0x10, 0)
#__________________________________ Functions _______________________________________#
#-----------------------------------------------#
def setProxy():
    """
    setProxy function to open chrome driver with new proxy after google_request_limit exceeded.
    """
    global urls_requested_Count,proxy,driver,chrome_options
    try:
        driver.close()
    except Exception as ex:
        ##logging.error('Except|setProxy(): driver quit - Exception: %s',ex)
        print('setProxy(): driver close - Exception: ', ex)
    finally:
        urls_requested_Count = 0
        curr_proxy_indx = 0
        check_proxy_lst()
        if len(VALID_PROXY_LST) > 0:
            try:
                try:
                    curr_proxy_indx = VALID_PROXY_LST.index(proxy)
                except ValueError:
                    curr_proxy_indx = -1
                except Exception as ex:
                    #logging.error('Except|setProxy(): get proxy index - Exception: %s', ex)
                    print(Fore.RED, 'setProxy() - get proxy index - Exception', ex, Fore.RESET)
                next_proxy_indx = curr_proxy_indx + 1 if (curr_proxy_indx + 1) < len(VALID_PROXY_LST) else 0
                proxy = VALID_PROXY_LST[next_proxy_indx]
                # --------- Set Proxy and chrome options ------#
                chrome_options = webdriver.ChromeOptions()
                chrome_options.add_argument('--proxy-server=%s' % proxy)
            except Exception as ex:
                # --------- Set Proxy and chrome options ------#
                #logging.error('Except|setProxy() - Exception: %s', ex)
                print(Fore.RED, 'setProxy(): - Exception', ex, Fore.RESET)
        else:
            chrome_options = None
        time.sleep(60)
        ##logging.error('Except|'_______setProxy()________')
        driver =  webdriver.Chrome(executable_path=chrome_path,chrome_options=chrome_options)
#------ Check if Proxy working ----------#
def check_proxy_lst():
    """
           check_proxy_lst function to check if all proxies in list working.
    """
    global VALID_PROXY_LST,driver,chrome_options
    try:
        VALID_PROXY_LST = []
        for prxy in PROXY_lst:
            err_text = ''
            try:
                # --------- Set Proxy and chrome options ------#
                chrome_options = webdriver.ChromeOptions()
                chrome_options.add_argument('--proxy-server=%s' % prxy)
                driver =  webdriver.Chrome(executable_path=chrome_path,chrome_options=chrome_options)
                driver.get("https://www.google.com/")
                time.sleep(5)
                proxy_err = driver.find_element(by=By.CSS_SELECTOR, value='#sub-frame-error-details')
                err_text = proxy_err.get_attribute('innerHTML')
            except Exception as ex:
                pass
            finally:
                chrome_options = None
                if 'wrong' and 'proxy' not in err_text:
                    VALID_PROXY_LST.append(prxy)
                try:
                    driver.close()
                except Exception as ex:
                    #logging.error('Except|check_proxy_lst(): driver close - Exception: %s', ex)
                    print('check_proxy_lst(): driver close - Exception  ', ex)
                time.sleep(10)
                #print(Fore.RED, "Func: check_proxy - if valid or not - Exception:  ", ex, Fore.RESET)
    except Exception as ex:
        #logging.error('Except|check_proxy_lst() - Exception: %s', ex)
        print(Fore.RED, "check_proxy_lst() - Exception:  ", ex, Fore.RESET)
    #finally:
        #logging.error('Except|'_______check_proxy_lst()________')
#------ Generate Full URL --------#
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
    elif url.startswith('/ - /'):
        Full_URL = scheme + '://' + Domain + url.replace('/ - /', '/')
    else:
        Full_URL = scheme + '://' + Domain + '/' + url
        #print("Stop")
    return Full_URL
#---------------------------------#
def download_PDF(pdf_url,part_num,optiontxt):
    pdf_downloaded = False
    file_path = '-'
    file_contentType = ''
    try:
        filename = pdf_url.split('/')[-1].split('.')[0]
        base_filename = part_num + '_' + optiontxt
        with urllib.request.urlopen(pdf_url) as response:
            info = response.info()
            file_contentType = info.get_content_type()
        if 'pdf' in str(file_contentType).lower():
            file_path = os.path.join(pdf_files_path, base_filename + ".pdf")
            urllib.request.urlretrieve(pdf_url, file_path)
        if os.path.isfile(file_path):
            pdf_downloaded = True
    except Exception as ex:
        print(Fore.RED, "download_PDF() - Exception:  ", ex, Fore.RESET)
    finally:
        return (pdf_downloaded,file_path)
#-------------------- Insert to file---------#
def insert_to_file(data_list):
    try:
        with open(parts_status_file, "a",encoding='utf-8') as resultFile:
            for data in data_list:
                data = str(data).replace('\n','').replace('\t','').strip()
                resultFile.write(data + '\t')
            resultFile.write('\n')
    except Exception as ex:
        print("insert_to_file() >>  Exception: ",ex)
#----------------------------------------------#
def start():
   try:
      main_url = 'https://www.norcomp.net/environmental-compliance/declaration-generator'
      driver.get(main_url)
   except BaseException as ex:
       print(">>> Request Url: ",str(ex))
   # ----
   try:
       for part_num in parts_lst:
           part_status = ''
           part_exception = ''
           part_compilance = '-'
           pdf_path = '-'
           try:
               print(">>>>>>> Part Number: ", str(part_num))
               compliance_ddl = Select(driver.find_element(by=By.NAME, value='comp-doc'))
               compliance_options = compliance_ddl.options
               input_part = driver.find_element(by=By.NAME, value='part')
               compilance_wanted = list(filter(lambda o: 'rohs' in str(o[-1].text).lower() or 'svhc' in str(o[-1].text).lower(),enumerate(compliance_options)))
               for option in compilance_wanted:
                   #--- Clear Compliance Type data -----#
                   compliance_ddl.select_by_index(0)
                   time.sleep(2)
                   #-----------------------------#
                   option_indx = option[0]
                   option_txt = str(option[-1].text).lower()
                   compliance_ddl.select_by_index(option_indx)
                   input_part.clear()
                   time.sleep(2)
                   input_part.send_keys(part_num)
                   time.sleep(3)
                   part_result =  driver.find_element(by=By.ID, value='autoResults')
                   result_childs = []
                   #------ If part_result has Children , means has add btn #
                   try:
                       result_childs = part_result.find_elements(by=By.XPATH, value='.//*')
                   except Exception as ex:
                       pass
                   if len(result_childs) > 0:
                       add_btn = part_result.find_element(by=By.TAG_NAME,value='a')
                       add_btn.click()
                       time.sleep(3)
                       #--------#
                       download_btn = driver.find_element(by=By.ID,value='dl-now')
                       download_lnk = download_btn.get_attribute('href')
                       #---- Get Full URL ----#
                       download_lnk = Generate_Full_URL(main_url,download_lnk)
                       pdf_downloaded, file_path = download_PDF(download_lnk,part_num,option_txt)
                       if pdf_downloaded:
                           pdf_path = file_path
                           if 'rohs' in option_txt:
                               part_compilance = 'ROHS'
                           elif 'svhc' in option_txt:
                               part_compilance = 'SVHC'
                           part_status = "3"
                           part_exception = "Yes"
                       else:
                           part_status = '0'
                           if 'rohs' in option_txt:
                               part_compilance = 'ROHS'
                               part_exception = 'No'
                           elif 'svhc' in option_txt:
                               part_compilance = 'SVHC'
                               part_exception = 'No'
                   else:
                       part_status = '0'
                       if 'rohs' in option_txt:
                           part_compilance = 'ROHS'
                           part_exception = 'No'
                       elif 'svhc' in option_txt:
                           part_compilance = 'SVHC'
                           part_exception = 'No'
                   try:
                       part_data = []
                       part_data.extend((part_num, part_status, part_exception, part_compilance, pdf_path))
                       insert_to_file(part_data)
                   except BaseException as ex:
                       print("____insert_to_file - Exception", str(ex))
                   print(part_compilance,"___", part_exception)
           except BaseException as ex:
               part_exception = str(ex)
               part_status = '-1'
               part_data = []
               part_data.extend((part_num, part_status, part_exception, part_compilance, pdf_path))
               insert_to_file(part_data)
               print("____Exception: ", str(ex))
           time.sleep(6)
   except BaseException as ex:
       print(">>> Loop over parts - Exception: ",str(ex))
#_____________________Start Execution _________________#
################################################################################################
if __name__ == '__main__':
    time.sleep(3)
    check_proxy_lst()
    if len(VALID_PROXY_LST) > 0:
        proxy = VALID_PROXY_LST[0]
        chrome_options.add_argument('--proxy-server=%s' % proxy)
    #------ Close Log File ------#
    if len(parts_lst) > 0:
        prefs = {"download.default_directory": pdf_files_path,  # IMPORTANT - ENDING SLASH V IMPORTANT
                 "directory_upgrade": True}
        chrome_options.add_experimental_option("prefs", prefs)
        driver = webdriver.Chrome(executable_path=chrome_path, chrome_options=chrome_options)
        # ------ Start--------#
        start()
        #---------Close log file and driver-------------#
        try:
            parts_status_file.close()
        except:
            pass
        try:
            driver.quit()
        except:
            pass
        #----------- Send Complete Msg -------#
        ctypes.windll.user32.MessageBoxW(None, "Please Check folder \n\n" + str(pdf_files_path) + "", "Completed",
                                         0x40, 0)
    else:
        ctypes.windll.user32.MessageBoxW(None, "No parts in Excel fle", "Completed",
                                         0x40, 0)


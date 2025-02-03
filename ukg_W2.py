from selenium import webdriver #Web driver activities
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC #Error handling
from selenium.webdriver.support.ui import WebDriverWait #Web driver wait
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import Select
from configparser import ConfigParser #Configuration read
from pathlib import Path #Path conversions
from os import listdir
from os.path import isfile, join
import os
import time
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
# import undetected_chromedriver as uc
import os #Path
import sys #System paths
import re
import csv
import ctypes #message box
import pyodbc
import logging #Activity logging
import time #Sleep
import pyautogui
import time
import pygetwindow
import shutil
import keyboard   #from webdriver_manager.chrome import ChromeDriverManager
 
 
download_dir = os.path.dirname(os.path.realpath(__file__))+'\\temp_downloads'
ConfigPath = os.path.dirname(os.path.realpath(__file__)) + '\\config.ini'
firefox_location='./'
 
 
global wait, driver, webpage_url, client_code, user_name, password, sq_1, sq_2, sq_3, sq_4, sq_5
global sql_server_name, sql_user_name, sql_password, sql_db, SQLconnection,min
 
 
def read_config_file():
    global webpage_url, client_code, user_name, password, sq_1, sq_2, sq_3, sq_4, sq_5
    global sql_server_name, sql_user_name, sql_password, sql_db
    try:
        # configuration entries
        config = ConfigParser()
        config.read(ConfigPath)
        # webpage_url = config.get ("Data", "webpage_URL")
        # user_name = config.get ("Data", "user_name")
        # password = config.get ("Data", "password")
        sql_server_name = config.get ("SQL", "server")
        sql_user_name = config.get ("SQL", "user")
        sql_password = config.get ("SQL", "password")
        sql_db = config.get ("SQL", "database")
        return(0)
    except Exception as e:
        print('config file read error')
        return(-1)
 
#--------------------------------------------------------------------------------------------------------
 
 
def setup():
    global firefox_location, download_dir, webpage_url, driver, wait
    binary = FirefoxBinary(firefox_location)
    profile = webdriver.FirefoxProfile()
    profile.set_preference("browser.download.folderList", 2)
    profile.set_preference("browser.download.manager.showWhenStarting", False)
    profile.set_preference("browser.download.dir", download_dir)
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/css, application/vnd.openxmlformats-officedocument.wordprocessingml.document, application/octet-stream, application/word, application/wordpad, image/png, image/bmp, image/jpeg, application/pdf, text/csv, text/html, text/plain, application/docx, application/zip")
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "image/jpeg")
    profile.set_preference("browser.helperApps.saveToDisk.image/jpeg", "application/octet-stream")
    profile.set_preference("browser.download.folderList", 2)
    profile.set_preference("browser.download.manager.showWhenStarting", False)
    profile.set_preference("browser.download.dir", download_dir)
    profile.set_preference("browser.download.useDownloadDir", True)
    profile.set_preference("browser.download.viewableInternally.enabledTypes", "")
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/css, application/vnd.openxmlformats-officedocument.wordprocessingml.document, application/octet-stream, application/word, application/wordpad, image/png, image/bmp, image/jpeg, application/pdf, text/csv, text/html, text/plain, application/docx, application/x-pdf, application/vnd.pdf, text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*")
    profile.set_preference("pdfjs.disabled", True)
    profile.set_preference("browser.download.manager.useWindow", False)
    profile.set_preference("browser.download.manager.closeWhenDone", True)
    profile.set_preference("print_printer", "Microsoft Print to PDF")
    profile.set_preference("print.always_print_silent", True)
    profile.set_preference("print.show_print_progress", False)
    profile.set_preference("print.save_as_pdf.links.enabled", True)
    profile.set_preference("browser.helperApps.alwaysAsk.force", False)
    profile.set_preference("plugin.disable_full_page_plugin_for_types", "text/css, application/vnd.openxmlformats-officedocument.wordprocessingml.document, application/octet-stream, application/word, application/wordpad, image/png, image/bmp,, image/jpeg, application/pdf, text/csv, text/html, text/plain, application/docx, application/x-pdf, application/vnd.pdf, text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*")
    # driver = webdriver.Firefox(service_log_path='NUL', firefox_profile=profile)
    driver = webdriver.Firefox()
    wait = WebDriverWait(driver, 10)
    driver.maximize_window()
    webpage_url = "https://nam04.safelinks.protection.outlook.com/?url=https%3A%2F%2Few45.ultipro.com%2F&data=05%7C02%7Cdocumenttransfers%40hcmunlocked.com%7Cd6c82143917142d5c2ac08dcb1863684%7Ced078563a25342048b117cf935797b6f%7C0%7C0%7C638580437238916509%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C0%7C%7C%7C&sdata=WAtO5XiqCRWxkV79zOWW3dMcR8gjIYxq33GFLvKULlo%3D&reserved=0"
    driver.get(webpage_url)
    time.sleep(1)
    wait = WebDriverWait(driver, 20)
    return 0
 
#--------------------------------------------------------------------------------------------------------
 
# def addToDatabase(emp_id,emp_name,category,document_title,each_file):
#     global sql_db
#     try:
#         cursor = SQLconnection.cursor()
#         insert_statement = """INSERT INTO [Studio_C].[dbo].[Emp_Files1] (
#                     [Src_Id],[Employee_ Name],[Category],[File_Name],[Document_Title],[IsDownloaded],[DownloadedDateTime])
#                     VALUES (?, ?, ?, ?, ?, ?, ?);"""
#         cursor.execute(insert_statement,(emp_id, emp_name,category,each_file,document_title,1,time.strftime('%Y-%m-%d %H:%M:%S')))
#         SQLconnection.commit()
#         print(document_title)
#         cursor.execute("UPDATE [Studio_C].[dbo].[Emp_Files] SET [File_Name] = ?, [IsDownloaded] = ?, [DownloadedDateTime] = ? WHERE [Src_Id] = ? AND [Document_Title] = ? ;",
#         (each_file,1, time.strftime('%Y-%m-%d %H:%M:%S'), emp_id,document_title))
#         SQLconnection.commit()
#     except:
#         print('insert to database error')

def addToDatabase(src_id,emp_name,document_title,category):
    try:
            cursor = SQLconnection.cursor()
            insert_statement = """
                INSERT INTO [Orchestra].[dbo].[Emp_Files_emp_docs] ([Employee_Number],[Emp_Name],[Category],[File_Name],[IsDownloaded]
                ) VALUES (?, ?, ?, ?, ?)
            """
            values = (src_id,emp_name,category,document_title, 1)
            cursor.execute(insert_statement, values)
            SQLconnection.commit()
            print(src_id," :Added to Emp_Files")

    except Exception as d:
        print('Database error which is ', d)
 
#--------------------------------------------------------------------------------------------------------
 
def remove_files():
    files_to_delete = os.listdir(download_dir)
    for filename in files_to_delete:
        file_path = os.path.join(download_dir, filename)
        os.remove(file_path)  # Delete the file
 
#--------------------------------------------------------------------------------------------------------
 
def createFolder(sFolder):
    isExist = os.path.exists(sFolder)
    if not isExist:
        os.makedirs(sFolder)
    return(0)
 
#--------------------------------------------------------------------------------------------------------
 
def login():
    global min
    #username input xpath
    # Find and interact with username field
    s=input("enter_before:")
    xpath_username = '//*[@id="ctl00_Content_Login1_UserName"]'
    element_username = driver.find_element(By.XPATH, xpath_username)
    element_username.click()
    element_username.clear()
    element_username.send_keys("documenttransfers@hcmunlocked.com")

    time.sleep(2)
 
    # Find and interact with password field
    xpath_password = '//*[@id="ctl00_Content_Login1_Password"]'
    element_password = driver.find_element(By.XPATH, xpath_password)
    element_password.click()
    element_password.clear()
    element_password.send_keys("Welcome@2024")
    time.sleep(2)

    # Find and click login button
    xpath_login_button = '//*[@id="ctl00_Content_Login1_LoginButton"]'
    element_login_button = driver.find_element(By.XPATH, xpath_login_button)
    element_login_button.click()
    # time.sleep(10)
    s= input("enter_after:")
 
    #Employee symbol
    xpath = '//*[@id="menu_admin"]'
    element = driver.find_element(By.XPATH,xpath)
    element.click()
    time.sleep(5)
    #My Employees
    xpath = '//*[@id="424"]'
    element = driver.find_element(By.XPATH,xpath)
    element.click()
    time.sleep(5)
   
    element = driver.find_element(By.ID,"ContentFrame")
    driver.switch_to.frame(element)
    element = driver.find_element(By.XPATH,'//*[@id="GridView1_firstSelect_0"]')
    select = Select(element)
    element.click()
    element.send_keys("Employee Number")
    time.sleep(2)
    element = driver.find_element(By.XPATH,'//*[@id="GridView1_Operator_0"]')
    element.send_keys("is")
    time.sleep(2)
 
#--------------------------------------------------------------------------------------------------------
 
# def createFolder(sFolder):
#     isExist = os.path.exists(sFolder)
#     if not isExist:
#         os.makedirs(sFolder)
#     return(0)
 
def connectSQL():
    global sql_server_name, sql_user_name, sql_password, sql_db, SQLconnection
    try:
        SQLconnection = pyodbc.connect('Driver={SQL Server};'
                            'Server=' + sql_server_name + ';'
                            'Database=' + sql_db + ';'
                            'UID=' + sql_user_name + ';'
                            'PWD=' + sql_password + ';'
                            'Trusted_Connection=no;')
        return(0)
    except Exception as e:
        print('SQL connection error')
        return(-1)
   
 
# Function to rename and move the file
# def rename_and_move_file(file_name, document_title, category, file_extension, target_dir, fld_name, download_dir,src_id,emp_name):
#     try:
#         # Define new folder path in the target directory based on category
#         new_folder = os.path.join(target_dir, category)
#         if not os.path.exists(new_folder):
#             os.makedirs(new_folder)
       
#         # Define new file path
#         new_file_path = os.path.join(new_folder, document_title + file_extension)
#         time.sleep(3)
       
#         # Rename and move the file
#         original_file = os.path.join(download_dir, file_name)
#         shutil.move(original_file, new_file_path)
#         print(f"Moved file '{file_name}' to '{new_file_path}'")

def rename_and_move_file(file_name, document_title, category, file_extension, target_dir, fld_name, download_dir, src_id, emp_name):
    try:
        # Define new folder path in the target directory based on category
        new_folder = os.path.join(target_dir, category)
        if not os.path.exists(new_folder):
            os.makedirs(new_folder)

        # Define the base new file path
        base_new_file_path = os.path.join(new_folder, document_title + file_extension)
        
        # Check if the file already exists and generate a new name if necessary
        new_file_path = base_new_file_path
        counter = 1
        while os.path.exists(new_file_path):
            new_file_path = os.path.join(new_folder, f"{document_title}_{counter}{file_extension}")
            counter += 1

        # Rename and move the file
        original_file = os.path.join(download_dir, file_name)
        shutil.move(original_file, new_file_path)
        print(f"Moved file '{file_name}' to '{new_file_path}'")

    # except Exception as e:
    #     print(f"Error: {e}")
        value = '"'+fld_name+'"'+","+'"'+document_title+'"'
        f= open('Berlin1_Downloaded_files.csv',"a")
        f.write(value)
        f.write('\n')
        f.close()
        addToDatabase(src_id,emp_name,document_title,category)
    except Exception as e:
        print(f"Error while renaming and moving the file: {e}")
 
def wait_for_download(download_dir, timeout=30):
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < timeout:
        time.sleep(1)
        dl_wait = any([filename.endswith('.part') for filename in os.listdir(download_dir)])
        seconds += 1
    return not dl_wait
   
#--------------------------------------------------------------------------------------------------------
 
def searchanddownload(fld_name, src_id, emp_name):
    global SQLconnection,min
    try:
        print("src Id ", src_id)
        driver.refresh()
        window_handles = driver.window_handles
        print("a ", len(window_handles))
        a = len(window_handles)
        b = 0
        while (a>1):
            driver.switch_to.window(window_handles[b-1])
            driver.close()
            b -=1
            a-=1
        window_handles = driver.window_handles
        driver.switch_to.window(window_handles[0])
        time.sleep(5)
        try:
            element = driver.find_element(By.ID,"ContentFrame")
            driver.switch_to.frame(element)
            time.sleep(3)
            print("switched")
        except:
            print("already switched")

        element = driver.find_element(By.XPATH,'//*[@id="GridView1_firstSelect_0"]')
        select = Select(element)
        element.click()
        element.send_keys("Employee Number")
        time.sleep(2)
        element = driver.find_element(By.XPATH,'//*[@id="GridView1_Operator_0"]')
        element.send_keys("is")
        time.sleep(2)
 
        input_element = driver.find_element(By.ID, 'GridView1_TextEntryFilterControlInputBox_0')
        input_element.click()
        input_element.clear()
        input_element.send_keys(src_id)
        time.sleep(3)
        keyboard.press_and_release('enter')
        time.sleep(4)
        # driver.refresh()
       
        try:
            element = driver.find_element(By.ID,"ContentFrame")
            driver.switch_to.frame(element)
        except:
            print('already in content frame \n')
        element = driver.find_element(By.XPATH, '//*[@id="ctl00_Content_GridView1"]/tbody/tr/td[2]')
        print(element.text)
        time.sleep(5)
 
        if element.text == src_id:
            #First employee
            element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_Content_GridView1"]/tbody/tr/td[1]/a'))
            )
            element.send_keys(Keys.ENTER)
       
            window_handles = driver.window_handles
            driver.switch_to.window(window_handles[1])
            time.sleep(4)
            driver.maximize_window()
            time.sleep(2)
            window_handles = driver.window_handles
            print("a ", len(window_handles))
            a = len(window_handles)
            b = 0
            while (a>2):
                driver.switch_to.window(window_handles[b-1])
                driver.close()
                b -=1
                a-=1
            window_handles = driver.window_handles
            driver.switch_to.window(window_handles[-1])
            # s = len(window_handles)
            # while (s>2):
            #     driver.switch_to.window(window_handles[b-1])
            #     driver.close()
       
            #Employee Documents
            # xpath = '//*[@id="1305"]'
            # element = driver.find_element(By.XPATH,xpath)
            # element.click()
            xpath = '//*[@id="1265"]'
            element = driver.find_element(By.XPATH,xpath)
            element.click()
            time.sleep(10)
 
            try:
                element = driver.find_element(By.ID, "ContentFrame")
                driver.switch_to.frame(element)
            except:
                print("No content frame found")
            time.sleep(3)

            try:
                xpath = '//*[@id="ctl00_Content_GridView1"]/tbody/tr/td/span'
                element = driver.find_element(By.XPATH,xpath)
                no_records = element.text
                print("no_records:",no_records)
                value = fld_name
                f= open('Berlin1_No records.csv',"a")
                f.write(value)
                f.write('\n')
                f.close()
            except:
                print("there are files")

            main_window = driver.current_window_handle
            # year_links = driver.find_elements(By.XPATH, "//a[contains(@href, 'javascript:pageLinkWithParams')]")
            # count= 0
            # requested_years = [2021,2022,2023,2024]
            # for i in range(len(year_links)):
            #     year_links = driver.find_elements(By.XPATH, "//a[contains(@href, 'javascript:pageLinkWithParams')]")
            #     year = year_links[i].text
            #     print("year:",year)
            #     requested_years = [2021,2022,2023,2024]
            #     if year in requested_years:
            #         count+=1
            #     else:
            #         break
            year_links = driver.find_elements(By.XPATH, "//a[contains(@href, 'javascript:pageLinkWithParams')]")
            count = 0
            requested_years = {2021, 2022, 2023, 2024}  # Using a set for faster lookup

            for link in year_links:
                year = int(link.text)  # Convert the text to an integer
                print("year:", year)
                
                if year in requested_years:
                    count += 1
                else:
                    break 
                    # continue
            print(count)
            count=str(count)
            years_count = len(year_links)
            print("year_counts:",years_count)
            print(f"Found {len(year_links)} rows.")
            years_count = str(len(year_links))
            value = '"'+fld_name+'"'+","+years_count+","+count
            f= open('Berlin1_Report_count.csv',"a")
            f.write(value)
            f.write('\n')
            f.close()

            for i in range(len(year_links)):
                try:
                    print("document:",i)
                    try:
                        element = driver.find_element(By.ID, "ContentFrame")
                        driver.switch_to.frame(element)
                    except:
                        print("No content frame found")
                    time.sleep(3)
                    
                    year_links = driver.find_elements(By.XPATH, "//a[contains(@href, 'javascript:pageLinkWithParams')]")
                    year = year_links[i].text
                    print("year:",year)
                    requested_years = {2021, 2022, 2023, 2024}
                    year_1 = int(year)
                    if year_1 in requested_years:
                        year_links[i].click()
                        # year = year_links[i].text
                        # print("year:",year)
                        time.sleep(2) 
                        print("year_clicking")

                        print_button = driver.find_element(By.ID, "ctl00_btnPrint")
                        print_button.click()
                        time.sleep(2)  # Give time for the new window to open
                        print("first_printing")
                        
                        # Switch to the new window (the print view)
                        # for handle in driver.window_handles:
                        #     if handle != main_window:
                        #         driver.switch_to.window(handle)
                        #         break
                        driver.switch_to.window(driver.window_handles[-1])
                        print("second_print_starting")
                        # try:
                        #     second_print_button = driver.find_element(By.XPATH, "//input[@class='printButton' and @value='print']")
                        #     second_print_button.click()
                        #     time.sleep(2)
                        #     print("second_print")
                        # except Exception as e:
                        #     print("error_clicking_second_print:",e)
                        try: 
                            keyboard.press_and_release('ctrl+p')
                            time.sleep(3)
                            keyboard.press_and_release('enter')
                            time.sleep(3)
                            pyautogui.typewrite("W2_"+year)
                            time.sleep(3)
                            keyboard.press_and_release('enter')
                            time.sleep(5)

                            createFolder("H:/Navya/downloads/Orchestra_W2_Term_New/Berlin/"+fld_name)
                            download_dir = r'C:\Users\RPATEAMADMIN\Downloads\\'
                            files = os.listdir(download_dir)
                            for file in files:
                                src=download_dir+file
                                dst="H:/Navya/downloads/Orchestra_W2_Term_New/Berlin/"+fld_name+'/'
                                shutil.move(src,dst)  
                            value = '"'+fld_name +'"'+","+year
                            print(value)
                                
                            f= open('Berlin1_Downloaded_files.csv',"a")
                            f.write(value)
                            f.write('\n')
                            f.close()
                            
                        except Exception as e:
                            print("error_clicking_enter:",e)
                            pyautogui.press('ctrl+p')
                            time.sleep(2)
                            print("trying")
                            pyautogui.press('enter')
                            # pyautogui.moveTo(x, y)
                            time.sleep(2)
                            pyautogui.typewrite("W2_"+year)
                            time.sleep(2)
                            pyautogui.press('enter')
                            time.sleep(2)
                            createFolder("H:/Navya/downloads/Orchestra_W2_Term_New/Berlin/"+fld_name)
                            download_dir = r'C:\Users\RPATEAMADMIN\Downloads\\'
                            files = os.listdir(download_dir)
                            for file in files:
                                src=download_dir+file
                                dst="H:/Navya/downloads/Orchestra_W2_Term_New/Berlin/"+fld_name+'/'
                                shutil.move(src,dst)  
                            value = '"'+fld_name +'"'+","+year
                            print(value)
                                
                            f= open('Berlin1_Downloaded_files.csv',"a")
                            f.write(value)
                            f.write('\n')
                            f.close()
                        
                        driver.close()

                        driver.switch_to.window(driver.window_handles[-1])
                        # try:
                        #     element = driver.find_element(By.ID,"ContentFrame")
                        #     driver.switch_to.frame(element)
                        #     time.sleep(3)
                        #     print("switched")
                        # except:
                        #     print("already switched")
                        try:
                            # xpath = '//*[@id="1265"]'
                            xpath ="//*[@aria-label='US Wage & Tax Statements']"
                            element = driver.find_element(By.XPATH,xpath)
                            element.click()
                        except:
                            driver.switch_to.window(driver.window_handles[-1])
                            try:
                                element = driver.find_element(By.ID,"ContentFrame")
                                driver.switch_to.frame(element)
                                time.sleep(3)
                                print("switched")
                            except:
                                print("already switched")
                                xpath ="//*[@aria-label='US Wage & Tax Statements']"
                                element = driver.find_element(By.XPATH,xpath)
                                element.click()
                    else:
                        break

                except Exception as e:
                    print(f"Error processing year link: {e}")
 
            # download_dir = r'C:\Users\RPATEAMADMIN\Downloads'
            # files = os.listdir(download_dir)
            # for file in files:
            #     src=download_dir+file
            #     dst='A:/Renuka/Projects/Orchestra/W2 docs/'+fld_name+'/'
            #     shutil.move(src,dst)  
            # value = fld_name +" :Downloaded"
            # print(value)
                
            # f= open('Downloaded_files.csv',"a")
            # f.write(value)
            # f.write('\n')
            # f.close()
 
            # target_dir = "A:/Renuka/Projects/Orchestra/W2 docs/"+fld_name+"/"

            # shutil.move(source_file, destination_file)
            

    
        return 0
 
    except Exception as e:
        print("Employee documents Not downloaded\n")
        f= open('Berlin1_Not_downloaded.csv',"a")
        f.write(fld_name)
        f.write('\n')
        f.close()
        print("error ",e)
        return -1
 
#--------------------------------------------------------------------------------------------------------
 
 
def main():
    global SQLconnection
    setup()
    read_config_file()
    connectSQL() 
    login()
 
    import csv
    rows = []
    with open("Berlin_term_roster.csv", 'r') as file:
        csvreader = csv.reader(file)
        #header = next(csvreader)
        for row in csvreader:
            rows.append(row)
    print(len(rows))
    # rows = []
    # select_statement = "SELECT [Employee_Number],[Last_Name],[First_Name],[IsDownloaded_Emp_docs]"
    # select_statement = select_statement + "FROM [Orchestra].[dbo].[Emp_List_Active]"
    # select_statement += "WHERE [Company] = 'BerlinRosen';"
    # cursor = SQLconnection.cursor()
    # cursor.execute(select_statement)
    # rows = cursor.fetchall()
    # cursor.close()
#266,267             #Term 448(Vm24)
    for i in range(422,len(rows)):
        each = rows[i]
        src_id = each[1].strip()
        print(src_id)
        src_id = src_id.zfill(6)
        # last_name = each[0].strip()
        # first_name = each[1].strip()
        emp_name = each[0].strip()
        # emp_name = last_name+", "+first_name
        print(emp_name)
        path_1 = r'C:\Users\RPATEAMADMIN\Downloads'
        files = os.listdir(path_1)
        # Iterate over the files and delete them
        for file in files:
            file_path = os.path.join(path_1, file)
            if os.path.isfile(file_path):
                os.remove(file_path)
                print(f"Deleted {file_path}")
        path = "H:/Navya/downloads/Orchestra_W2_Term_New/Berlin/"
        directory_contents = os.listdir(path)
        fld_name = emp_name+" ("+src_id+")"
        if fld_name not in directory_contents:
            print("entered_searchanddownload")
            # err_f = 1
            # count_flag = 1
            # while(err_f):
            res = searchanddownload(fld_name, src_id, emp_name)
            print("res=",res)
            # if res == 0:
            #     cursor = SQLconnection.cursor()
            #     cursor.execute("UPDATE [Orchestra].[dbo].[Emp_List_Active] SET [IsDownloaded_Emp_docs] = ? WHERE [Employee_Number] = ?;",
            #     (1, src_id))
            #     print(src_id," :Updated in Emp_List")
            #     SQLconnection.commit()


            # elif res<0:
            #     print('Employee error '+src_id+" "+emp_name)
            #     continue

 
#--------------------------------------------------------------------------------------------------------
 
main()
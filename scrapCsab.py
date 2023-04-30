# import required modules
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
import os
import pandas as pd





opt = Options()
opt.add_argument('--disable-blink-features=AutomationControlled')
opt.add_argument('--start-maximized')
# headless mode
# opt.add_argument('--headless')
# opt.add_extension('IDM-Integration-Module.crx')
driver = webdriver.Chrome(executable_path='chromedriver.exe', options=opt)
driver.get('http://josaa.admissions.nic.in/Applicant/seatallotmentresult/currentorcr.aspx')

# wait till the page loads
try:    
    myElem = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@type='submit']")))
    print("Page is ready!")
    
    for i in range(2,7):
        select=driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlroundno_chosen")
        time.sleep(5)
        select.click()
        time.sleep(1)
        option = driver.find_element(By.XPATH, f"//li[@data-option-array-index='{i}']")
        option.click()
        time.sleep(1)
        InstituteSelect = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlInstype_chosen")
        time.sleep(1)
        InstituteSelect.click()
        time.sleep(1)
        option = driver.find_element(By.XPATH, "//li[@data-option-array-index='1']")
        option.click()
        time.sleep(1)
        instituteNameSelect = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlInstitute_chosen")
        time.sleep(1)
        instituteNameSelect.click()
        time.sleep(1)
        option = driver.find_element(By.XPATH, "//li[@data-option-array-index='1']")
        option.click()
        time.sleep(1)
        academicType = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlBranch_chosen")
        time.sleep(1)
        academicType.click()
        time.sleep(1)
        option = driver.find_element(By.XPATH, "//li[@data-option-array-index='1']")
        option.click()
        time.sleep(1)
        seatType = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlSeattype_chosen")
        time.sleep(1)
        seatType.click()
        time.sleep(1)
        option = driver.find_element(By.XPATH, "//li[@data-option-array-index='1']")
        option.click()
        time.sleep(1)

        button = driver.find_element(By.XPATH, "//input[@type='submit']")
        button.click()
        time.sleep(10)
        myDriver =WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//tr[@class='bg-secondary text-white']")))
        print("Page is ready!")
        allRows = driver.find_elements(By.XPATH, "//tr")
        print(len(allRows))
        allRow =[]
        counter = 0
        for row in allRows:
            counter = counter + 1
            print(f"{counter}/{len(allRows)} ")
            # get all child elements of a row
            cells = row.find_elements(By.XPATH, "./*")
            # print all cells
            currentRow = []
            for cell in cells:
                currentRow.append(cell.text)
            allRow.append(currentRow)
        
        pd.DataFrame(allRow).to_excel(f"output_josa_round_{i}.xlsx")



                
except TimeoutException:
    print ("Loading took too much time!")

# def checkLoginState():   # Check if user is logged in
#     try:
#         driver.find_element(By.XPATH,"//a[@href='/login']")
        
#         isPresent = True;
#     except:
#         isPresent = False;
#     return isPresent;

# def Login():                # Login to sleepdata.org
#     driver.find_element(By.XPATH,"//input[@id='user_email']").send_keys(myId)
#     driver.find_element(By.XPATH,"//input[@id='user_password']").send_keys(myPassword)
#     driver.find_element(By.XPATH,"//input[@type='submit']").click()
# print(checkLoginState())
# #calling Login function if user is not logged in to sleepdata.org
# if(checkLoginState()==False):
#     print("Already logged in")
# else:
#     driver.find_element(By.XPATH,"//a[@href='/login']").click()
#     try:
#         myElem = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//div[@class="sign-up-form-title"]')))
#         Login()
#         print("Page is ready!")
#     except TimeoutException:
#         print ("Loading took too much time!")
#     print("Logged in")
# #function to check if file is already downloaded
# def checkFile():
#     pathOfVisit1="D:\\DSP\\mros\\polysomnography\\annotations-events-profusion\\visit1"
#     fileList=os.listdir(pathOfVisit1)
#     return fileList
# #function to download file
# def logLinks():
#     for i in range(1,31):
#         driver.get('https://sleepdata.org/datasets/mros/files/polysomnography/annotations-events-profusion/visit1?page='+str(i))
#         time.sleep(1)
#         links = driver.find_elements(By.XPATH,"//a[@data-turbolinks='false']")
#         listFile=checkFile()
#         for link in links:
#             fileName=link.get_attribute('innerHTML')
#             if(fileName in listFile):
#                 print("File already downloaded")
#                 pass
#             else:
#                 print("Downloading file: "+fileName)
#                 try:
#                     link.click()   
#                 except:
#                     try:
#                         link.click()
#                         time.sleep(1)
#                     except:
#                         pass

# logLinks()
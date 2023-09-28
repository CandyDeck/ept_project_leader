from selenium import webdriver 
from selenium.webdriver import Firefox
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


import time 
import pandas as pd
import zipfile



from datetime import datetime
import os
import re
import shutil



today = datetime.now()
isExist = os.path.exists(str(today.strftime('%Y%m%d')))
if not isExist:
   # Create a new directory because it does not exist
   os.mkdir(today.strftime('%Y%m%d'))
   print("The new directory is created!")
else:
    print('exist already')


profile = webdriver.FirefoxProfile()
profile.set_preference("browser.download.folderList", 2)
profile.set_preference("browser.download.manager.showWhenStarting", False)
profile.set_preference("browser.download.dir", os.getcwd() + "\\" + str(today.strftime('%Y%m%d')) + "\\"  )
profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/x-gzip")
driver = webdriver.Firefox(firefox_profile=profile)

driver.get('https://login.dfo.no/mga/sps/authsvc?PolicyId=urn:ibm:security:authentication:asf:dfolandingpage')
driver.implicitly_wait(10) # wait until page will be loaded

ddelement= driver.find_element(By.ID, "select-identityProvider").click()
feide = driver.find_element(By.XPATH,  "//*[text()='Feide']").click()
time.sleep(2)
button = driver.find_element(By.XPATH,'/html/body/div/div/section/div/div/div/form/fieldset[2]/button')
time.sleep(2)
button.click()
affiliation = driver.find_element(By.ID,"org_selector_filter")
affiliation.send_keys("ntnu")
ntnu = driver.find_element(By.XPATH,'/html/body/div/article/section[2]/div[1]/form/div[1]/ul/li[20]/div').click()
button2 = driver.find_element(By.ID, "selectorg_button")
button2.click()

username=driver.find_element(By.ID,"username")
password = driver.find_element(By.ID,"password")
username.send_keys("***")
password.send_keys("***")
button3 = driver.find_element(By.XPATH,'/html/body/div/article/section[2]/div[1]/form[1]/button')
button3.click()

window_before = driver.window_handles[0]

unit4 = driver.find_element(By.XPATH,'/html/body/div/div/div/section/div/div/div/div/div/div/ul/li[2]/a')
unit4.click()
time.sleep(10)
window_after = driver.window_handles[1]
driver.switch_to.window(window_after)


prosjektstyring = driver.find_element(By.XPATH,"//*[text()='Prosjektstyring']")
driver.execute_script("arguments[0].click();", prosjektstyring)

#prosjektstyring.click()

time.sleep(5)
expander = driver.find_elements(By.XPATH, '//div[contains(@class, "u4-menu-folder-expander")]')
delt=driver.find_element(By.ID, str(expander[0].get_attribute('id')))
delt.click()
intern=driver.find_element(By.XPATH,"//*[text()='Intern prosjektrapport']")
intern.click()


project_number = pd.read_csv(os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))+ "\\" + 'project_number_bidrag_oppdrag.csv')

correspondance = pd.DataFrame(columns=['project_number','file_number'])
for a in project_number['0']:
    print(a)
    time.sleep(2)
    iframe = driver.find_element(By.XPATH, "//iframe[contains(@id, 'ext-gen')]")
    wait = WebDriverWait(driver, 10)
    wait.until(EC.frame_to_be_available_and_switch_to_it(iframe))
    frame2 = driver.find_element(By.XPATH, "//frame[contains(@id, 'contentContainerFrame')]")
    wait.until(EC.frame_to_be_available_and_switch_to_it(frame2))
    time.sleep(1)

    project = driver.find_element(By.ID, "b_s4_l1s4_ctl00_project_i")
    time.sleep(1)
    project.send_keys(str(a))
    time.sleep(1)
    period = driver.find_element(By.ID, "b_s4_l1s4_ctl00_period_i")
    period.click()
    
    wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Lagre']")))
    lagre = driver.find_element(By.XPATH,"//span[text()='Lagre']") 
    time.sleep(1)
    driver.execute_script("arguments[0].click();", lagre)
    time.sleep(1)

    driver.switch_to.window(window_after)
    time.sleep(1)
    message = driver.find_elements(By.XPATH,"//div[contains(@id, 'u4_messageoverlay_success-')]")
    file_number = re.findall(r'\d+',message[6].text)[0]  
    print(file_number)
    ok =driver.find_element(By.XPATH,"//*[contains(@id,'u4_pagebutton-')]") 
    time.sleep(1)
    
    d={'project_number':[a],'file_number': [file_number]}
    df = pd.DataFrame(data=d)
    
    correspondance=pd.concat([correspondance, df])

    driver.execute_script("arguments[0].click();", ok)
        
    time.sleep(2)


# select_all_sheets = driver.find_element(By.XPATH,"//div[contains(@class,'HeaderCheckButton')]")
correspondance.to_csv( os.getcwd() + "\\" + str(today.strftime('%Y%m%d')) + "\\"+ 'correspondance_bidrag_oppdrag_'+ str(today.strftime('%Y%m%d')) +'.csv',index = False)


iframe = driver.find_element(By.XPATH, "//iframe[contains(@id, 'ext-gen')]")
wait = WebDriverWait(driver, 50)
wait.until(EC.frame_to_be_available_and_switch_to_it(iframe))
frame2 = driver.find_element(By.XPATH, "//frame[contains(@id, 'contentContainerFrame')]")
wait.until(EC.frame_to_be_available_and_switch_to_it(frame2))
time.sleep(1)
rapport = driver.find_element(By.XPATH,"//*[text()='Dine bestilte rapporter']")
rapport.click()
time.sleep(5)
driver.switch_to.window(window_after)


iframe2 = driver.find_elements(By.XPATH, "//iframe[contains(@id, 'ext-gen')]")
wait.until(EC.frame_to_be_available_and_switch_to_it(iframe2[1]))
frame3 = driver.find_element(By.XPATH, "//frame[contains(@id, 'contentContainerFrame')]")
wait.until(EC.frame_to_be_available_and_switch_to_it(frame3))
select_all_sheets = driver.find_element(By.XPATH,"//div[contains(@class,'HeaderCheckButton')]")
select_all_sheets.click()
# box =driver.find_element(By.XPATH, "//th[contains(@class, 'GridCell NotMovable NotResizable ColumnCheckbox')]")
# box.click()

download_button = driver.find_element(By.XPATH,"//*[text()='Last ned']")
download_button.click()
time.sleep(30)

time.sleep(5)
driver.close()
driver.quit()


'''Rename zip file downloaded from unit4'''

# folder path
dir_path = os.getcwd() + "\\" + str(today.strftime('%Y%m%d')) 


# Iterate directory
for file in os.listdir(dir_path):
    # check only zip files
    if file.endswith('.zip'):
        print(file)
        old_name = dir_path + "\\" + file
        print(old_name)
        new_name = dir_path + "\\" + str(today.strftime('%Y%m%d')) + '.zip'
        os.rename(old_name, new_name)

isExist = os.path.exists(dir_path + "\\" + 'excel_project_sheets')
if not isExist:
   # Create a new directory because it does not exist
   os.mkdir(dir_path + "\\" + 'excel_project_sheets')
   
'''Extract files from zip file and store in folder we just created'''

with zipfile.ZipFile(new_name, "r") as zf:
        zf.extractall(path=dir_path + "\\" + 'excel_project_sheets')
os.remove(new_name)

project_number = pd.read_csv(os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))+ "\\" +'project_number_bidrag_oppdrag.csv')

#project_number_bidrag = pd.read_csv(os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))+ "\\" +'project_number_bidrag.csv')
list_project =  pd.read_excel(os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))+ "\\" + str(today.strftime('%Y%m%d')) + '_list_project.xlsx','WORKORDERS')
correspondance = pd.read_csv(os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))+ "\\" +'correspondance_bidrag_oppdrag_' + str(today.strftime('%Y%m%d')) + '.csv')
#project_number_bidrag = pd.read_csv(os.getcwd() + "\\20230502\\" +'project_number_bidrag.csv')
# list_project =  pd.read_excel(os.getcwd() + "\\20230427\\" + str('20230427') + '_list_project.xlsx','WORKORDERS')
# correspondance = pd.read_csv(os.getcwd() + "\\20230427\\" +'correspondance_' + str('20230427') + '.csv')


for a in list_project['Prosjektleder (T)'].unique() :
    os.mkdir(os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))+ "\\"+str(a)+"\\Project_reports\\")



for row in list_project.index:
    if (list_project.loc[row,['Protype (T)']].to_string(header=False, index=False))=='Bidrag':
        project_num = int(list_project.loc[row,['Prosjektnr.']].to_string(header=False, index=False))
        project_PL =  list_project.loc[row,['Prosjektleder (T)']].to_string(header=False, index=False)
        if project_num in correspondance['project_number'].values:
            excel_num = int(correspondance.loc[correspondance['project_number']==project_num,'file_number'].to_string(header=False, index=False))
            print(project_num,project_PL,excel_num)
            shutil.copy(os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))+ "\\excel_project_sheets\\poerapa_" + str(excel_num)+'.xlsx', os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))+"\\" + str(project_PL)+"\\Project_reports"+"\\poerapa_" + str(excel_num)+'.xlsx')
        else :
            continue
for row in list_project.index:
    if (list_project.loc[row,['Protype (T)']].to_string(header=False, index=False))=='Oppdrag':
        project_num = int(list_project.loc[row,['Prosjektnr.']].to_string(header=False, index=False))
        project_PL =  list_project.loc[row,['Prosjektleder (T)']].to_string(header=False, index=False)
        if project_num in correspondance['project_number'].values:
            excel_num = int(correspondance.loc[correspondance['project_number']==project_num,'file_number'].to_string(header=False, index=False))
            print(project_num,project_PL,excel_num)
            shutil.copy(os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))+ "\\excel_project_sheets\\poerapa_" + str(excel_num)+'.xlsx', os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))+"\\" + str(project_PL)+"\\Project_reports"+"\\poerapa_" + str(excel_num)+'.xlsx')
        else :
            continue


'''RENAME EXCEL SCHEET'''


for a in list_project['Prosjektleder (T)'].unique() :
    
    dir_path = os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))  + "\\" + a +"\\Project_reports"
    for file in os.listdir(dir_path):
        if  file.startswith('poerapa'):
            old_name = dir_path + "\\" + file
            project_name = pd.read_excel(old_name,'_Intern prosjektrapport')
            name = project_name.loc[project_name['Unnamed: 1']=='Prosjekt navn',['Unnamed: 2']].to_string(header=False, index=False)
            name = name.replace(':', '')
            new_name = dir_path + "\\" + name + '.xlsx'
            if os.path.exists(new_name):
                
                new_name = dir_path + "\\" + name + '2.xlsx'

                os.rename(old_name, new_name)
                continue
            os.rename(old_name, new_name)

'''EXCEL SHEETS ARE RENAMED. WE NEED TO CREATE A EXCEL FILE WITH NAME AND BUTTON LINKING TO PROJECT BUDGET'''
 
'''CREATION OF PAGE 1'''

for a in list_project['Prosjektleder (T)'].unique() :
    #os.mkdir(os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))+ "\\"+str(a)+"\\Budget\\")

    page2=pd.DataFrame(columns=['Project_name','Link to project budget'])
    dir_path = os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))  + "\\" + a 
    dir_path_budget = os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))  + "\\" + a + "\\Project_reports"
    

            
    for file in os.listdir(dir_path_budget):
        if  file.endswith('.xlsx'):
            d={'Project_name':[file.split('.x')[0]],'Link to project budget': ['=HYPERLINK("' + 'Project_reports\\' +  file +  '","click here")']}
            df = pd.DataFrame(data=d)
            page2=pd.concat([page2, df])
    page2.to_excel(dir_path + '\\project_reports.xlsx',index=False,encoding='utf-8-sig')
          
            
  

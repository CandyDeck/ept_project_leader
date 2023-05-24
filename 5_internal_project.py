from selenium import webdriver 

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time 

from datetime import datetime
import os
import pandas as pd 
import numpy as np
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
username.send_keys("candyd")
password.send_keys("20SimGroup18")
button3 = driver.find_element(By.XPATH,'/html/body/div/article/section[2]/div[1]/form[1]/button')
button3.click()

window_before = driver.window_handles[0]

unit4 = driver.find_element(By.XPATH,'/html/body/div/div/div/section/div/div/div/div/div/div/ul/li[2]/a')
unit4.click()
time.sleep(10)
window_after = driver.window_handles[1]
driver.switch_to.window(window_after)

rapporter = driver.find_element(By.XPATH,"//*[text()='Rapporter']")
driver.execute_script("arguments[0].click();", rapporter)

spørringer_regnskap = driver.find_element(By.XPATH,"//*[text()='Spørringer regnskap']")
driver.execute_script("arguments[0].click();", spørringer_regnskap)


spørring_bilag = driver.find_element(By.XPATH,"//*[text()='Spørring bilag - dim 1-7']")
driver.execute_script("arguments[0].click();", spørring_bilag)

iframe = driver.find_element(By.XPATH, "//iframe[contains(@id, 'ext-gen')]")
wait = WebDriverWait(driver, 10)
wait.until(EC.frame_to_be_available_and_switch_to_it(iframe))
time.sleep(5)
frame2 = driver.find_element(By.XPATH, "//frame[contains(@id, 'contentContainerFrame')]")
wait.until(EC.frame_to_be_available_and_switch_to_it(frame2))
time.sleep(5)

koststed =driver.find_element(By.XPATH,'//input[contains(@id, "_dim_1=_i")]')
koststed.clear()
koststed.send_keys(64250500)
time.sleep(2)

period =driver.find_element(By.XPATH,'//input[contains(@id, "_period<>_i")]')
period.clear()
period.send_keys(str(today.year)+str(0)+str(0))


søk = driver.find_element(By.XPATH,"//*[text()='Søk']")
driver.execute_script("arguments[0].click();", søk)

time.sleep(180)

eksport = driver.find_element(By.XPATH,"//*[text()='Eksport']")
eksport.click()

save= driver.find_element(By.XPATH,"//*[text()='Default [.xlsx]']")
save.click()
time.sleep(20)


old_name = os.getcwd() + "\\" + str(today.strftime('%Y%m%d')) + "\\" + "Spørring bilag - dim 1-7.xlsx"
print(old_name)
new_name = os.getcwd() + "\\" + str(today.strftime('%Y%m%d')) + "\\" + str(today.strftime('%Y%m%d')) + "_spørring_bilag.xlsx"
os.rename(old_name, new_name)

driver.close()
driver.quit()

#list_project =  pd.read_excel(os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))+ "\\" + str(today.strftime('%Y%m%d')) + '_list_project.xlsx','WORKORDERS')
list_project =  pd.read_excel(os.getcwd() + "\\" + str(20230522)+ "\\" + str(20230522) + '_list_project.xlsx','WORKORDERS')

for a in list_project['Prosjektleder (T)'].unique() :
    os.mkdir(os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))+ "\\"+str(a)+"\\Internal_projects\\")

os.mkdir(os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))+ "\\1_Other_Internal_projects\\")


'''il faut maintenant garder que les drifts et separer par project number'''
#sporring = pd.read_excel(os.getcwd() + "\\" + str(today.strftime('%Y%m%d')) + "\\" + str(today.strftime('%Y%m%d')) + "_spørring_bilag.xlsx",'Spørring bilag - dim 1-7')
sporring = pd.read_excel(os.getcwd() + "\\" + str(20230523) + "\\" + str(20230523) + "_spørring_bilag.xlsx",'Spørring bilag - dim 1-7')
sporring_new = sporring.copy()

sporring_new = sporring_new[sporring_new['Protype (T)']==("Intern")]

sporring_new= sporring_new.reset_index(drop=True)
for a in sporring_new.index:
    
    if sporring_new.loc[a,'Konto']==2160:
        sporring_new.loc[a,'Konto (T)']='Saldo'

for a in reversed(range(len(sporring_new))):
    
    konto = sporring_new.loc[a,'Konto']
    if konto<2160:
        print(konto,a,sporring_new.index[a])
        sporring_new = sporring_new.drop([sporring_new.index[a]])
    if  (konto  in [3953,6040,6050,3951,6001]):
        sporring_new = sporring_new.drop([sporring_new.index[a]])
        
sporring_new= sporring_new.reset_index(drop=True)

   # if str(sporring_new.loc[a,'Konto']).startswith('5'):
   #     sporring_new.loc[a,'Konto (T)']='Salary'
   #     sporring_new.loc[a,'Konto']=5000
        
for num in sporring_new['Delprosjekt'].unique():
    print(num)
    if num in list_project['Arbeidsordre-ID'].unique():
        project_name = list_project.loc[list_project['Arbeidsordre-ID']==num,['Navn på arbeidsordre']].to_string(header=False,index=False)
        project_leader = list_project.loc[list_project['Arbeidsordre-ID']==num,['Prosjektleder (T)']].to_string(header=False,index=False)
        project_name=project_name.replace('*','')
        project_name=project_name.replace('/','')
        sporring_new.loc[(sporring_new['Konto']==2160)&(sporring_new['Delprosjekt']==num),['Anlegg/Ansattnr (T)']]=project_leader
              #for num in sporring_new['Delprosjekt'].unique():
              #    if num == 984323129:
           
        pl = sporring_new.loc[sporring_new['Delprosjekt']==num]
        for a in pl.index:
            if(pd.isnull(pl.loc[a,'Anlegg/Ansattnr (T)'])):
                pl.loc[a,'Anlegg/Ansattnr (T)']='other'
        project_name = pl.loc[(pl.index[0]),'Delprosjekt (T)']
        project_name=project_name.replace('*','')   
        project_name=project_name.replace('/','')
        pl_pivot = pl.pivot_table(values='Beløp', index=['Konto','Konto (T)','Anlegg/Ansattnr (T)'], columns='Periode', aggfunc='sum')
        if (str(pl_pivot. columns[-1]) == str(today.strftime('%Y%m'))):
            del pl_pivot[pl_pivot. columns[-1]]
        
        col=[]
        for a in pl_pivot.columns:
            if str(a).startswith('2'):
                col.append(a)

        pl_pivot['Total'] = pl_pivot[col].sum(axis=1)
        pl_pivot['a enlever'] = ''
        pl_pivot['NEW SALDO'] = ''

        pl_pivot.loc[(pl_pivot.index[0]),'NEW SALDO']=pl_pivot['Total'].sum(axis=0)
        pl_pivot.rename(columns = {"a enlever" : ""}, inplace = True)
 
        #pl_pivot.to_excel(os.getcwd() + "\\" + str(today.strftime('%Y%m%d')) + "\\"  + str(project_leader) + "\\Internal_projects\\" + str(project_name)+'.xlsx',encoding='utf-8-sig')
        pl_pivot.to_excel(os.getcwd() + "\\" + str(20230523) + "\\"  + str(project_leader) + "\\Internal_projects\\" + str(project_name)+'.xlsx',encoding='utf-8-sig')

    else :
        project_name = sporring_new.loc[(sporring_new['Delprosjekt']==num)&(sporring_new['Konto']==2160),['Delprosjekt (T)']].to_string(header=False,index=False)
        
        project_name=project_name.replace('*','')
        project_name=project_name.replace('/','')

        project_leader = 'Other'
        sporring_new.loc[(sporring_new['Konto']==2160)&(sporring_new['Delprosjekt']==num),['Anlegg/Ansattnr (T)']]=project_leader
               
        pl = sporring_new.loc[sporring_new['Delprosjekt']==num]
        for a in pl.index:
            if(pd.isnull(pl.loc[a,'Anlegg/Ansattnr (T)'])):
                pl.loc[a,'Anlegg/Ansattnr (T)']='other'
        project_name = pl.loc[(pl.index[0]),'Delprosjekt (T)']
        project_name=project_name.replace('*','')
        project_name=project_name.replace('/','')


        pl_pivot = pl.pivot_table(values='Beløp', index=['Konto','Konto (T)','Anlegg/Ansattnr (T)'], columns='Periode', aggfunc='sum')
        col=[]
        for a in pl_pivot.columns:
            if str(a).startswith('2'):
                col.append(a)
        pl_pivot['Total'] = pl_pivot[col].sum(axis=1)
        pl_pivot['a enlever'] = ''
        pl_pivot['NEW SALDO'] = ''

        pl_pivot.loc[(pl_pivot.index[0]),'NEW SALDO']=pl_pivot['Total'].sum(axis=0)
        pl_pivot['Total']=''
        pl_pivot.rename(columns = {"a enlever" : ""}, inplace = True)
     
        #pl_pivot.to_excel(os.getcwd() + "\\" + str(today.strftime('%Y%m%d')) + "\\1_Other_Internal_projects\\" + str(project_name)+'.xlsx',encoding='utf-8-sig')
        pl_pivot.to_excel(os.getcwd() + "\\" + str(20230523) + "\\1_Other_Internal_projects\\" + str(project_name)+'.xlsx',encoding='utf-8-sig')
        
        
        
    
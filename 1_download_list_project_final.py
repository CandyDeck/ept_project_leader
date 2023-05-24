from selenium import webdriver 

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains

import time 

from datetime import datetime
import os

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


prosjektstyring = driver.find_element(By.XPATH,"//*[text()='Prosjektstyring']")
driver.execute_script("arguments[0].click();", prosjektstyring)

#prosjektstyring.click()

time.sleep(5)

ny_rapport = driver.find_element(By.XPATH,"//*[text()='Opprett ny rapport']")
ny_rapport.click()
time.sleep(3)


while(True) : 
     try : 
         element = driver.find_elements(By.XPATH, '//input[contains(@id, "u4_textfield-") and contains(@placeholder, "Søk")]')

         break
     except : 
      pass
  
time.sleep(5)
input_id = str(element[1].get_attribute('id'))
time.sleep(1)
input_id_space=driver.find_element(By.ID,input_id)
time.sleep(1)
input_id_space.send_keys('Arbeidsordre')
time.sleep(2)
action = ActionChains(driver)
action.send_keys(Keys.ENTER)
time.sleep(3)
driver.find_element(By.XPATH, '//*[@class="u4-menu-search-result u4f-menu-search-enquiries-result"]').click()
time.sleep(3)


kriterium =driver.find_element(By.XPATH,"//*[text()='Legg til kriterium']")
driver.execute_script("arguments[0].click();", kriterium)
mer =driver.find_element(By.XPATH,"//span[text()='Mer...']")
mer.click()


item_tree_view_kostnadssted = driver.find_element(By.XPATH,"//tr[contains(@data-qtip,'Kostnadssted')]")
checkbox_kostnadssted = driver.find_elements(By.XPATH,"//tr[contains(@data-qtip,'Kostnadssted')]//td[contains(@id,'ext-gen')]")
checkbox_kostnadssted[1].click()



ok_button = driver.find_element(By.XPATH,'//a[contains(@id, "u4_pagebutton-") and contains(@class, "x-btn u4-pagebutton u4-standard-button-small u4-standard-button-happy x-unselectable x-box-item x-btn-u4-standard-button-small x-noicon x-btn-noicon x-btn-u4-standard-button-small-noicon")]')
ok_button.click()
time.sleep(3)
spaces_precis =driver.find_elements(By.XPATH,'//div[contains(@id, "u4_form-") and contains(@class,"x-box-target")]//div[contains(@id, "u4f_ib_attValueListCriterionView-")]//input[contains(@id, "u4_typeahead-") and contains(@class,"x-form-field x-form-text u4-form-field-input u4-form-field-trigger-input")]')
space_precis_id = str(spaces_precis[4].get_attribute('id'))
space_precis_id_select=driver.find_element(By.ID,space_precis_id)
space_precis_id_select.send_keys(64250500)
resultat =driver.find_element(By.XPATH,"//*[text()='Vis resultat']") 
driver.execute_script("arguments[0].click();", resultat)
time.sleep(5)
button_add_column = driver.find_element(By.XPATH,'//div[contains(@class, "u4-tablecustomisation")]')
button_add_column.click()


item_tree_view_prosjektleder = driver.find_elements(By.XPATH,"//tr[contains(@data-qtip,'Prosjektleder')]")
checkbox_prosjektleder = driver.find_elements(By.XPATH,"//tr[contains(@data-qtip,'Prosjektleder')]//td[contains(@id,'ext-gen')]")
checkbox_prosjektleder[6].click()


item_tree_view_kundenavn = driver.find_elements(By.XPATH,"//tr[contains(@data-qtip,'Kundenr')]")
checkbox_kundenavn = driver.find_elements(By.XPATH,"//tr[contains(@data-qtip,'Kundenr')]//td[contains(@id,'ext-gen')]")
checkbox_kundenavn[10].click()

item_tree_view_project = driver.find_elements(By.XPATH,"//tr[contains(@data-qtip,'Prosjekt')]")
checkbox_project = driver.find_elements(By.XPATH,"//tr[contains(@data-qtip,'Prosjekt')]//td[contains(@id,'ext-gen')]")
time.sleep(5)
driver.execute_script("arguments[0].click();", checkbox_project[27])

item_tree_view_project_type = driver.find_elements(By.XPATH,"//tr[contains(@data-qtip,'Protype')]")
checkbox_project_type = driver.find_elements(By.XPATH,"//tr[contains(@data-qtip,'Protype')]//td[contains(@id,'ext-gen')]")
checkbox_project_type[2].click()




ok_button2= driver.find_element(By.XPATH,'//a[contains(@id, "u4_pagebutton-") and contains(@class, "x-btn u4-pagebutton u4-standard-button-small u4-standard-button-happy x-unselectable x-btn-toolbar x-box-item x-toolbar-item x-btn-u4-standard-button-toolbar-small x-noicon x-btn-noicon x-btn-u4-standard-button-toolbar-small-noicon")]')
ok_button2.click()
time.sleep(5)
resultat =driver.find_element(By.XPATH,"//*[text()='Vis resultat']") 
driver.execute_script("arguments[0].click();", resultat)

time.sleep(10)

lagre =driver.find_element(By.XPATH,"//*[text()='Lagre']") 
driver.execute_script("arguments[0].click();", lagre)
time.sleep(2)

navn = driver.find_element(By.XPATH,"//*[text()='Navn']")
navn_id=str(navn.get_attribute('id'))
label = navn_id.split('-label')[0]
navn_file=str(label)+str('-inputEl')
navn_file_space=driver.find_element(By.ID,navn_file)
navn_file_space.send_keys(str(today.strftime('%Y%m%d')))
#navn_file_space.send_keys('test1005_2')

time.sleep(3)

privat_folder = driver.find_element(By.XPATH,"//div[contains(@title,'Privat')]")
privat_folder.click()

lagre2 =driver.find_elements(By.XPATH,"//span[text()='Lagre']") 

driver.execute_script("arguments[0].click();", lagre2[1])
time.sleep(0)



while(True) : 
     try : 
         ok3 = driver.find_element(By.XPATH,'//a[contains(@id, "u4_pagebutton-") and contains(@class,"x-btn u4-pagebutton u4-standard-button-small u4-standard-button-happy x-unselectable x-btn-u4-standard-button-small x-noicon x-btn-noicon x-btn-u4-standard-button-small-noicon")]')


         break
     except : 
      pass
  
time.sleep(2)    

# wait = WebDriverWait(driver, 10)
# wait.until(EC.element_to_be_clickable(ok3))
ok3.click()

eksport = driver.find_element(By.XPATH,"//span[text()='Eksport']")
driver.execute_script("arguments[0].click();", eksport)


file_save = driver.find_element(By.XPATH,'//a[contains(@data-qtip,"Run Browser [.xlsx]")]')
file_save.click()
time.sleep(5)

old_name = os.getcwd() + "\\" + str(today.strftime('%Y%m%d')) + "\\" + str(today.strftime('%Y%m%d')) + "_Browser.xlsx"
print(old_name)
new_name = os.getcwd() + "\\" + str(today.strftime('%Y%m%d')) + "\\" + str(today.strftime('%Y%m%d')) + "_list_project.xlsx"
os.rename(old_name, new_name)

cross = driver.find_element(By.XPATH,'//div[contains(@id,"reportengine_exportview") and contains(@class,"u4-floatingcontainer-background-close-icon")]')
driver.close()
driver.quit()

## Download the list of Project from unit4

The user of the script should be aware that his username and password should replace the *** :

```python
username=driver.find_element(By.ID,"username")
password = driver.find_element(By.ID,"password")
username.send_keys("***")
password.send_keys("***")
```

The user should also be aware that the right kostnadsted should be written in full, as a successsion of 8 digit or, it is also possible to enter the 4 first digit and complete the search with *:

```python
spaces_precis =driver.find_elements(By.XPATH,'//div[contains(@id, "u4_form-") and contains(@class,"x-box-target")]//div[contains(@id, "u4f_ib_attValueListCriterionView-")]//input[contains(@id, "u4_typeahead-") and contains(@class,"x-form-field x-form-text u4-form-field-input u4-form-field-trigger-input")]')
space_precis_id = str(spaces_precis[4].get_attribute('id'))
space_precis_id_select=driver.find_element(By.ID,space_precis_id)
space_precis_id_select.send_keys('********') '''6425*'''
```

The user has to specified a name for the file he wants to download :
```python
navn_file_space.send_keys(str(today.strftime('%Y%m%d')))
```

The user has also the possibility to rename the file:

```python
old_name = os.getcwd() + "\\" + str(today.strftime('%Y%m%d')) + "\\" + str(today.strftime('%Y%m%d')) + "_Browser.xlsx"
print(old_name)
new_name = os.getcwd() + "\\" + str(today.strftime('%Y%m%d')) + "\\" + str(today.strftime('%Y%m%d')) + "_list_project.xlsx"
os.rename(old_name, new_name)
```

## 2.  Create, for each project leader, a directory

The list of employees is available from a Teams' channel.
We import the list and store it in a database : 
```python
employees = pd.read_excel('C:\\Users\\candyd\\NTNU\\EPT HR og Økonomi - General\\Oversikter\\NY_Oversikt over ansatte og utlysninger.xlsx','Ansatte')
```
We also convert the list of projects (obtained from point 1) into a dataframe : 
```python
project = pd.read_excel(os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))+ "\\" + str(today.strftime('%Y%m%d')) + '_list_project.xlsx','WORKORDERS')
```
For each project leader appearing in the list (Prosjekleder (T)) in the list we create, in the folder whose name is today's date, a directory named under the project leader. In these project leaders directories, we also copy a list of project belonging to them. This individual list of project is named like : **"name of the project leader"_project_overview.xlsx** :
```python
for a in project['Prosjektleder (T)'].unique():
    isExist = os.path.exists(os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))+ "\\"+ str(a))
    if not isExist:
        # Create a new directory because it does not exist
        os.mkdir(os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))+ "\\"+ str(a))
        print("The new directory is created!")
    else:
        print('exist already')



    pl = project.loc[project['Prosjektleder (T)']==a]
    print(a,pl)
    pl['Dato fra']= pl['Dato fra'].dt.floor('T')
    pl['Dato til']= pl['Dato til'].dt.floor('T')

    pl.to_excel(str(today.strftime('%Y%m%d'))+"/"+str(a) + '/'+ str(a)+'_project_overview.xlsx',index=False,encoding='utf-8-sig')
```
Then, from the list of projects, we create a new list (**six_digit_project_number**) containing all the 6 digit project numbers related project types **Bidrag** and **Oppdrag**. This list is then saved under **project_number_bidrag_oppdrag.csv** :
```python
for row in project.index:
    if (project.loc[row,['Protype (T)']].to_string(header=False, index=False))=='Bidrag':
        if not project.loc[row,['Prosjektnr.']].to_string(header=False, index=False) in six_digit_project_number:
            six_digit_project_number.append(project.loc[row,['Prosjektnr.']].to_string(header=False, index=False))

for row in project.index:
    if (project.loc[row,['Protype (T)']].to_string(header=False, index=False))=='Oppdrag':
        if not project.loc[row,['Prosjektnr.']].to_string(header=False, index=False) in six_digit_project_number:
            six_digit_project_number.append(project.loc[row,['Prosjektnr.']].to_string(header=False, index=False))
            
df = pd.DataFrame(six_digit_project_number)

df.to_csv(os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))+ "\\" +'project_number_bidrag_oppdrag.csv',index=False)
```

Finally, from the list of employees, we create a file called "ansatte_'+ str(pl_folder[0])+'.xlsx'"

## 3. Downloaf the excel sheets from unit4 (externally founding projects)

From unit4, we download individual excel sheets containing the budget. 
This is done using selenium.
From point2, we have the list of project numbers.
As we download an excel file form Unit4, a number is allocated to this file. This number having nothing to do with the project number, we need to keep track for this number and the project number.
This is done simply by creating a new dataframe called **correspondance**, in which the 2 columns are named **project_number** and **file_number**:
```python
correspondance = pd.DataFrame(columns=['project_number','file_number'])
```
When all files are stored, we download from all the excel sheets at once from Unit4. These are stored in a .zip file which we rename
```python
for file in os.listdir(dir_path):
    # check only zip files
    if file.endswith('.zip'):
        print(file)
        old_name = dir_path + "\\" + file
        print(old_name)
        new_name = dir_path + "\\" + str(today.strftime('%Y%m%d')) + '.zip'
        os.rename(old_name, new_name)
```

Then, we unzip this .zip file and store all excel sheets in a folder called **excel_project_sheets**
Finally, knowing the correspondance between the project numbers and the file numbers, we can rename the excel sheets with the name of the correponding project.
These excel sheets are finally stored in the right project leader repository


## 4. Create budget files for each internally founded projects. 


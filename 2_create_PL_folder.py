import pandas as pd
from datetime import datetime
import os
import math
today = datetime.now()

employees = pd.read_excel(os.getcwd() + "\\" + 'NY_Oversikt over ansatte og utlysninger.xlsx','Ansatte')
employees = employees[['Navn', 'Stillingskode','Tittel','Startdato','Sluttdato','Statsborgerskap','Ansettelsesramme','Ansettelsesforhold','Prosjektledernr.','Veileder','HR-ansvarlig','Prosjektnummer']]

project = pd.read_excel(os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))+ "\\" + str(today.strftime('%Y%m%d')) + '_list_project.xlsx','WORKORDERS')
project_number = []



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
    pl.to_excel(str(today.strftime('%Y%m%d'))+"/"+str(a) + '/'+ str(a)+'_project_overview.xlsx',index=False,encoding='utf-8-sig')


six_digit_project_number =[]
#project_list = pd.read_excel(os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))+ "\\" + str(today.strftime('%Y%m%d')) + '_list_project.xlsx','WORKORDERS')
for row in project.index:
    if (project.loc[row,['Protype (T)']].to_string(header=False, index=False))=='Bidrag':
        if not project.loc[row,['Prosjektnr.']].to_string(header=False, index=False) in six_digit_project_number:
            six_digit_project_number.append(project.loc[row,['Prosjektnr.']].to_string(header=False, index=False))

df = pd.DataFrame(six_digit_project_number)

df.to_csv(os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))+ "\\" +'project_number_bidrag.csv',index=False)
                                                             

for row in employees.index:
    pl_num = employees.loc[row,'Prosjektledernr.']
    if (len(str(pl_num)))>8:
        part_a = pl_num.rsplit('/',1)[0]
        part_b = pl_num.rsplit('/',1)[1]
        
        pl_a = (project.loc[project['Prosjektleder'] == int(part_a[1:]), 'Prosjektleder (T)'].unique())[0]
        pl_b = (project.loc[project['Prosjektleder'] == int(part_b[1:]), 'Prosjektleder (T)'].unique())[0]
        
        
        new_data_a = pd.DataFrame([employees.iloc[202].values], index=[len(employees)], columns=employees.columns)  
        new_data_a['Veileder'] = pl_a
        new_data_a['Prosjektledernr.']  = part_a
        employees=pd.concat([employees, new_data_a])

        new_data_b = pd.DataFrame([employees.iloc[202].values], index=[len(employees)], columns=employees.columns)  
        new_data_b['Veileder'] = pl_b
        new_data_b['Prosjektledernr.']  = part_b
        employees=pd.concat([employees, new_data_b])
        employees.drop([row], axis=0, inplace=True)


for project_leader in employees['Prosjektledernr.'].unique(): 
    if not pd.isna(project_leader):
        #print(project_leader)
        page2 = employees.    loc[employees['Prosjektledernr.']==project_leader]    
        print(project_leader,str(project_leader)[1:])
        #print(page2)
        pl_folder = (project.loc[project['Prosjektleder'] == int(str(project_leader)[1:]), 'Prosjektleder (T)'].unique())
        if(pl_folder.size > 0):
            print(pl_folder[0])
            page2.to_excel(str(today.strftime('%Y%m%d'))+"/"+str(pl_folder[0]) + '/ansatte_'+ str(pl_folder[0])+'.xlsx',index=False,encoding='utf-8-sig')


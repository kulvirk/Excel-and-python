
import xlrd
import pandas as pd
import numpy as np
from  numpy import matrix
import datetime
import openpyxl
from openpyxl import load_workbook
import xlsxwriter

report_in2= pd.read_excel('report.xls')


report_in2.rename(columns={'Prioty':'Priority','Asgndgrp':'Assigned_group'},inplace=True)
report_in=report_in2.drop(columns=['State','Date'])
report_in=report_in.dropna(axis=0)

#output filename
report_out=pd.read_excel('new_report.xlsx')

#new column to add
new_column=((report_out.shape)[1])

# new team list
new_teamlist=[]

# All types of Priorities
Priority_values=[x for x in ((report_in['Priority'].unique()).tolist()) if str(x) != 'nan']


# In[3]:


#Combining teams
st1=['A','E','D','C','B','Dsupport', 'G','BSupport']

nd2=['Esup','ESupport','STeam']

rd3=['Eupt','Eort']

th4=['3rd Team']

th5=['GI Support']

th6=['FSupport','HSupport','TASManagement']

th7=['Dport','Et','M/Team','Main','7thSupport']




report_in['Assigned_group'].replace(st1,'st1_new',inplace=True)
report_in['Assigned_group'].replace(nd2,'nd2_new',inplace=True)
report_in['Assigned_group'].replace(rd3,'rd3_new',inplace=True)
report_in['Assigned_group'].replace(th4,'th4_new',inplace=True)
report_in['Assigned_group'].replace(th5,'th5_new',inplace=True)
report_in['Assigned_group'].replace(th6,'th6_new',inplace=True)
report_in['Assigned_group'].replace(th7,'th7_new',inplace=True)

#Total teams
Total_teams=report_in['Assigned_group'].nunique()
Team_names=(report_in['Assigned_group'].unique()).tolist()    
#report_in
#team_index_pri=report_out.index[report_out['Metrics of incident arrival'].str.contains('Priority',na=False)].tolist()


# In[4]:


def append_unique(l, val): 
    if val not in l: 
        l.append(val)


# In[5]:


def findindexstr(n,name):
    team_index_all=report_out.index[report_out['column with name of report'].str.contains(name,na=False)].tolist()
    team_index=team_index_all[n-1]
    print(name,team_index)
    return  team_index


# In[6]:


def findindex(n,team_name):
    team_index_all=report_out.index[report_out['column with name of report']==team_name].tolist()
    #print(team_index_all)
    if len(team_index_all)==0:
        append_unique(new_teamlist,team_name)
        #print(new_teamlist)
    else:
        team_index=team_index_all[n-1]
        print(team_name,team_index)
        return  team_index 


# In[7]:


table=report_in.groupby(['Priority','Assigned_group']).size().unstack()
table2=report_in.groupby(report_in['Priority']).size()
P_all=[]
for P in (1.0,2.0,3.0,4.0):
    try:
        P_1=table2.loc[P]
        P_all.append(P_1)
    except:
        P_1=0
        P_all.append(P_1)
print (P_all)        

def values(P_n,team_name):
    try:
        value=int(table.loc[P_n,team_name])
    except:
        value=0
    return value


# In[8]:


def enter_value(index,C_n,value):
    report_out.at[index,C_n]=value


# In[14]:


def update_outfile():
    #df=pd.DataFrame(['Metrics of incident arrival'])
                                 
    df2 = report_out
    writer = pd.ExcelWriter('new_report.xlsx',engine='xlsxwriter',datetime_format='dd/mm/yy')
    
    df2.to_excel(writer, header=None,sheet_name='sheet1',index=False,startrow=1)
    workbook=writer.book
    header_fmt = workbook.add_format({"bg_color": "#FADBD8",'border':1, 'bold': True})
    header_fmt2 = workbook.add_format({'font_name': 'Arial Black','font_size': 12, 'bold': True})
    date = datetime.datetime.strptime('01-03-01', "%d-%m-%y")
    worksheet=writer.sheets['sheet1']
    worksheet.write('A1','Metrics of incident arrival',header_fmt2)
    worksheet.conditional_format('A2:Z1000',{'type':'date',
                               'criteria':'greater than',
                               'value':date,
                               'format':header_fmt,
                            'multi_range': 'A2:Z2 A13:Z13 A24:Z24 A35:Z35'})
    border_format=workbook.add_format({
                            'border':1,
                            'align':'left',
                            'font_size':16
                           })
    
    worksheet.conditional_format('A1:Z1000',{'type' : 'no_blanks' , 'format' : border_format} )
    

    writer.save()


# In[15]:


now = datetime.datetime.now()
todaysDate=now.strftime("%d/%m/%y")
for P in range(1,5):
    Pri='Priority'
    Ttl='Total'
    I_p=findindexstr(P,Pri)
    I_t=findindexstr(P,Ttl)
    V_p=todaysDate
    V_t=P_all[P-1]
    C_p=new_column
    enter_value(I_p,C_p,V_p)
    enter_value(I_t,C_p,V_t)
    
    
    


# In[16]:


for T_n in Team_names:
    for P in range(1,5):
        I_n=findindex(P,T_n)
        V_n=0
        C_n=new_column
        enter_value(I_n,C_n,V_n)
update_outfile()


# In[17]:


for P_n in Priority_values:
    for T_n in Team_names:
        #print (int(P_n),str(T_n))
        I_n=findindex(int(P_n),T_n)
        V_n=values(P_n,T_n)
        C_n=new_column
        enter_value(I_n,C_n,V_n)
update_outfile()     



                                    

print("CERAGON Weekly Utilization REPORT STARTED")

import pandas as pd
import numpy as np
import paramiko 
import datetime
from datetime import datetime, timedelta
import win32com.client as wincl
from openpyxl import load_workbook
import os
import glob
import re
import warnings
warnings.filterwarnings('ignore')


da7=(datetime.now() - timedelta(1)).strftime('%Y%m%d')

dm7=(datetime.now() - timedelta(1)).strftime('%d_%m_%Y')

da=(datetime.now() - timedelta(1)).strftime('%d_%m_%Y')
do=(datetime.now() - timedelta(1)).strftime('%Y%m%d')
dn=(datetime.now() - timedelta(1)).strftime('%d-%m-%Y')

M=(datetime.now() - timedelta(1)).strftime('%d%m%Y')
M1=(datetime.now() - timedelta(2)).strftime('%d%m%Y')
M2=(datetime.now() - timedelta(3)).strftime('%d%m%Y')
M3=(datetime.now() - timedelta(4)).strftime('%d%m%Y')
M4=(datetime.now() - timedelta(5)).strftime('%d%m%Y')
M5=(datetime.now() - timedelta(6)).strftime('%d%m%Y')
M6=(datetime.now() - timedelta(7)).strftime('%d%m%Y')

print(M)
print(M1)
print(M2)
print(M3)
print(M4)
print(M5)
print(M6)

# For PM Files
DA = (datetime.now() - timedelta(0)).strftime('%d-%b-%y')

DA0= datetime.now() - timedelta(0)
DA1= datetime.now() - timedelta(1)
DA2= datetime.now() - timedelta(2)
DA3 = datetime.now() - timedelta(3)
DA4 = datetime.now() - timedelta(4)
DA5= datetime.now() - timedelta(5)
DA6= datetime.now() - timedelta(6)
DA7= datetime.now() - timedelta(7)


DA7 = datetime.strftime(DA7, '%d-%b-%y')
DA6 = datetime.strftime(DA6, '%d-%b-%y')
DA5 = datetime.strftime(DA5, '%d-%b-%y')
DA4 = datetime.strftime(DA4, '%d-%b-%y')
DA3 = datetime.strftime(DA3, '%d-%b-%y')
DA2 = datetime.strftime(DA2, '%d-%b-%y')
DA1 = datetime.strftime(DA1, '%d-%b-%y')
DA0 = datetime.strftime(DA0, '%d-%b-%y')

# For MYCOM Files


today = datetime.now()

seven_days_ago = today - timedelta(days=7)

# Check if any day in the last 7 days has a day less than 10
use_e_format = any((seven_days_ago + timedelta(days=i)).day < 10 for i in range(7))

# Choose the format based on whether any day < 10
if use_e_format:
    date_format = '%e %b %y'  # No leading zero if any day is less than 10
else:
    date_format = '%d %b %y'  # With leading zero if all days >= 10

# Generate and print the formatted dates for the last 7 days
formatted_dates = [(seven_days_ago + timedelta(days=i)).strftime(date_format) for i in range(7)]

# Assign to variables (MYCOM8 to MYCOM0)

MYCOM8, MYCOM7, MYCOM6, MYCOM5, MYCOM4, MYCOM3, MYCOM2= formatted_dates




##################### Reading files######################################

d1 = pd.read_excel(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera Daily Utilization\OUTPUT\TXN_VIL_CERAGON_MW_Link_Utilization_Daily_'+M+'.xlsx',sheet_name= 'Total_Backup')
d2 = pd.read_excel(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera Daily Utilization\OUTPUT\TXN_VIL_CERAGON_MW_Link_Utilization_Daily_'+M1+'.xlsx',sheet_name='Total_Backup')
d3 = pd.read_excel(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera Daily Utilization\OUTPUT\TXN_VIL_CERAGON_MW_Link_Utilization_Daily_'+M2+'.xlsx',sheet_name='Total_Backup')
d4 = pd.read_excel(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera Daily Utilization\OUTPUT\TXN_VIL_CERAGON_MW_Link_Utilization_Daily_'+M3+'.xlsx',sheet_name='Total_Backup')
d5 = pd.read_excel(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera Daily Utilization\OUTPUT\TXN_VIL_CERAGON_MW_Link_Utilization_Daily_'+M4+'.xlsx',sheet_name='Total_Backup')
d6 = pd.read_excel(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera Daily Utilization\OUTPUT\TXN_VIL_CERAGON_MW_Link_Utilization_Daily_'+M5+'.xlsx',sheet_name='Total_Backup')
d7 = pd.read_excel(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera Daily Utilization\OUTPUT\TXN_VIL_CERAGON_MW_Link_Utilization_Daily_'+M6+'.xlsx',sheet_name='Total_Backup')

print ( " Files Read " )


d1_Uniq=d1[['Server','Circle','LT','Link Configuration','uniq link','Site A Name','Site B Name','Site A IP','Site B IP','Site A Physical Port','Site B Physical Port','Site A Radio','Site B Radio','Site A Short name',
              'Site B Short name','CHANNEL SPACING','MAX_E1','Modulation']]
d2_Uniq=d2[['Server','Circle','LT','Link Configuration','uniq link','Site A Name','Site B Name','Site A IP','Site B IP','Site A Physical Port','Site B Physical Port','Site A Radio','Site B Radio','Site A Short name',
              'Site B Short name','CHANNEL SPACING','MAX_E1','Modulation']]
d3_Uniq=d3[['Server','Circle','LT','Link Configuration','uniq link','Site A Name','Site B Name','Site A IP','Site B IP','Site A Physical Port','Site B Physical Port','Site A Radio','Site B Radio','Site A Short name',
              'Site B Short name','CHANNEL SPACING','MAX_E1','Modulation']]
d4_Uniq=d4[['Server','Circle','LT','Link Configuration','uniq link','Site A Name','Site B Name','Site A IP','Site B IP','Site A Physical Port','Site B Physical Port','Site A Radio','Site B Radio','Site A Short name',
              'Site B Short name','CHANNEL SPACING','MAX_E1','Modulation']]
d5_Uniq=d5[['Server','Circle','LT','Link Configuration','uniq link','Site A Name','Site B Name','Site A IP','Site B IP','Site A Physical Port','Site B Physical Port','Site A Radio','Site B Radio','Site A Short name',
              'Site B Short name','CHANNEL SPACING','MAX_E1','Modulation']]
d6_Uniq=d6[['Server','Circle','LT','Link Configuration','uniq link','Site A Name','Site B Name','Site A IP','Site B IP','Site A Physical Port','Site B Physical Port','Site A Radio','Site B Radio','Site A Short name',
              'Site B Short name','CHANNEL SPACING','MAX_E1','Modulation']]
d7_Uniq=d7[['Server','Circle','LT','Link Configuration','uniq link','Site A Name','Site B Name','Site A IP','Site B IP','Site A Physical Port','Site B Physical Port','Site A Radio','Site B Radio','Site A Short name',
              'Site B Short name','CHANNEL SPACING','MAX_E1','Modulation']]

df=pd.concat([d1_Uniq,d2_Uniq,d3_Uniq,d4_Uniq,d5_Uniq,d6_Uniq,d7_Uniq])


# KeeP ROB

def custom_sort(val):
    if val == 'ROB':
        return 0  # Priority for ROB
    elif val == 'BIH':
        return 1  # Keep BIH after ROB
    else:
        return 2

df_sorted = df.sort_values(by='Server', key=lambda col: col.map(custom_sort))
df = df_sorted.drop_duplicates(subset='uniq link', keep='first')


print ( " Rename Column Name ")

d11=d1[['uniq link','MAX_TP']]
d12=d2[['uniq link','MAX_TP']]
d13=d3[['uniq link','MAX_TP']]
d14=d4[['uniq link','MAX_TP']]
d15=d5[['uniq link','MAX_TP']]
d16=d6[['uniq link','MAX_TP']]
d17=d7[['uniq link','MAX_TP']]

#d11=d11.rename(columns={'MAX_TP':'MAX_TP1'}).drop_duplicates(subset='uniq link', keep='first') 
#d12=d12.rename(columns={'MAX_TP':'MAX_TP2'}).drop_duplicates(subset='uniq link', keep='first') 
#d13=d13.rename(columns={'MAX_TP':'MAX_TP3'}).drop_duplicates(subset='uniq link', keep='first') 
#d14=d14.rename(columns={'MAX_TP':'MAX_TP4'}).drop_duplicates(subset='uniq link', keep='first') 
#d15=d15.rename(columns={'MAX_TP':'MAX_TP5'}).drop_duplicates(subset='uniq link', keep='first') 
#d16=d16.rename(columns={'MAX_TP':'MAX_TP6'}).drop_duplicates(subset='uniq link', keep='first') 
#d17=d17.rename(columns={'MAX_TP':'MAX_TP7'}).drop_duplicates(subset='uniq link', keep='first') 

d11=d11.rename(columns={'MAX_TP':'MAX_TP1'})
d12=d12.rename(columns={'MAX_TP':'MAX_TP2'})
d13=d13.rename(columns={'MAX_TP':'MAX_TP3'})
d14=d14.rename(columns={'MAX_TP':'MAX_TP4'})
d15=d15.rename(columns={'MAX_TP':'MAX_TP5'})
d16=d16.rename(columns={'MAX_TP':'MAX_TP6'})
d17=d17.rename(columns={'MAX_TP':'MAX_TP7'})

df['Site A Name']=df['Site A Name'].str.strip()
df['Site B Name']=df['Site B Name'].str.strip()


print ( " Read Combined Pm File " )

# ** Find Capacity **


Cdf = df.loc[df['Site A Radio']!='RFU-50E']

#Capacity_order = {'2QAM':20,'4QAM':40,'8QAM':60,'16QAM':85,'32QAM':108,'64QAM':134,'128QAM':159,'256QAM':181,'512QAM':194,'1024QAM':215,'1024QAMLight':226,
#                  '2048QAM':240,'32QAM':108,'4096QAM':300}


Capacity_order = {'QPSK':42,'2QAM':20,'4QAM':40,'8QAM':60,'16QAM':85,'32QAM':108,'64QAM':134,'128QAM':159,'256QAM':181,'512QAM':194,'1024QAMLight':215,'1024QAM':215,
                  '2048QAM':240,'4096QAM':300}
           	           
Cdf['Capacity']=Cdf['Modulation'].map(Capacity_order)

def calculate_capacity(row):    
    if "Xpic" in row['LT']:  
        if row['CHANNEL SPACING'] == "56MHz":
            return row['Capacity'] * 4 
        else:
            return row['Capacity'] * 2          
    elif row['LT'] == "1+0" and row['CHANNEL SPACING'] == "56MHz":
        return row['Capacity'] * 2        
    return row['Capacity']

# Apply the function to the DataFrame
Cdf['Total_Capacity'] = Cdf.apply(calculate_capacity, axis=1)



## --------------------------------------------------------------------------------------------------------

# Only 50CX
Cdf50=df.loc[df['Site A Radio']=='RFU-50E']

Capacity_order = {'BPSK9':46,'BPSK10':92,'BPSK':186,'QPSK':373,'8QAM':576,'16QAM':768,'32QAM':960,'64QAM':1153,'128QAM':1346,'256QAM':1538,'512QAM':1730}
              	           
Cdf50['Capacity']=Cdf50['Modulation'].map(Capacity_order)

def calculate_capacity(row):    
    if "Xpic" in row['LT']:  
        if row['CHANNEL SPACING'] == "56MHz":
            return row['Capacity'] * 4 
        else:
            return row['Capacity'] * 2          
    elif row['LT'] == "1+0" and row['CHANNEL SPACING'] == "56MHz":
        return row['Capacity'] * 2        
    return row['Capacity']

# Apply the function to the DataFrame
Cdf50['Total_Capacity'] = Cdf50.apply(calculate_capacity, axis=1)


## --------------------------------------------------------------------------------------------------------

Comb = pd.concat([Cdf,Cdf50])
df = Comb.copy()



# Merging Start ---------------------

df=pd.merge(df,d11,on='uniq link',how='left') # Merge 1
df=pd.merge(df,d12,on='uniq link',how='left') # Merge 2
df=pd.merge(df,d13,on='uniq link',how='left') # Merge 3
df=pd.merge(df,d14,on='uniq link',how='left') # Merge 4
df=pd.merge(df,d15,on='uniq link',how='left') # Merge 5
df=pd.merge(df,d16,on='uniq link',how='left') # Merge 6
df=pd.merge(df,d17,on='uniq link',how='left') # Merge 7


# Merging Done  ---------------------

df['UTIL_D1']=round(df['MAX_TP1']/df['Total_Capacity']*100,2)
df[['UTIL_D1']] =df[['UTIL_D1']].clip(upper=100)

df['UTIL_D2']=round(df['MAX_TP2']/df['Total_Capacity']*100,2)
df[['UTIL_D2']] =df[['UTIL_D2']].clip(upper=100)

df['UTIL_D3']=round(df['MAX_TP3']/df['Total_Capacity']*100,2)
df[['UTIL_D3']] =df[['UTIL_D3']].clip(upper=100)

df['UTIL_D4']=round(df['MAX_TP4']/df['Total_Capacity']*100,2)
df[['UTIL_D4']] =df[['UTIL_D4']].clip(upper=100)

df['UTIL_D5']=round(df['MAX_TP5']/df['Total_Capacity']*100,2)
df[['UTIL_D5']] =df[['UTIL_D5']].clip(upper=100)

df['UTIL_D6']=round(df['MAX_TP6']/df['Total_Capacity']*100,2)
df[['UTIL_D6']] =df[['UTIL_D6']].clip(upper=100)

df['UTIL_D7']=round(df['MAX_TP7']/df['Total_Capacity']*100,2)
df[['UTIL_D7']] =df[['UTIL_D7']].clip(upper=100)


#dff = dff.sort_values(by='Utilization', ascending=False)

# Count>90
col_cera=['UTIL_D1', 'UTIL_D2', 'UTIL_D3', 'UTIL_D4', 'UTIL_D5', 'UTIL_D6', 'UTIL_D7']

df_count_90=df[col_cera]

cera1 = df_count_90.iloc[:,0] > 90
cera2 = df_count_90.iloc[:,1] > 90
cera3 = df_count_90.iloc[:,2] > 90
cera4 = df_count_90.iloc[:,3] > 90
cera5 = df_count_90.iloc[:,4] > 90
cera6 = df_count_90.iloc[:,5] > 90
cera7 = df_count_90.iloc[:,6] > 90

cera_final = pd.concat([cera1,cera2,cera3,cera4,cera5,cera6,cera7],axis=1)

df['count>90']=cera_final.sum(axis=1)

# Count>100
df_count_100=df[col_cera]

cera8 = df_count_100.iloc[:,0] >= 100
cera9 = df_count_100.iloc[:,1] >= 100
cera10= df_count_100.iloc[:,2] >= 100
cera11= df_count_100.iloc[:,3] >= 100
cera12= df_count_100.iloc[:,4] >= 100
cera13= df_count_100.iloc[:,5] >= 100
cera14= df_count_100.iloc[:,6] >= 100

cera_final = pd.concat([cera8,cera9,cera10,cera11,cera12,cera13,cera14],axis=1)

df['count>100']=cera_final.sum(axis=1)

# Remove blanks from Utilization
Final_df =df.dropna(subset=['UTIL_D1', 'UTIL_D2', 'UTIL_D3', 'UTIL_D4', 'UTIL_D5', 'UTIL_D6', 'UTIL_D7'],how='all')

# Remove Duplicates Logic
Final_df['UTIL_SUM'] = Final_df[['UTIL_D1', 'UTIL_D2', 'UTIL_D3', 'UTIL_D4', 'UTIL_D5', 'UTIL_D6', 'UTIL_D7']].sum(axis=1)
Final_df = Final_df.sort_values(by='UTIL_SUM', ascending=False).reset_index(drop=True).drop_duplicates(subset='uniq link', keep='first')
Final_df.drop(columns=['UTIL_SUM'],inplace=True)


print (" Pivot Start " )

col_ops=list(['Circle','LT','uniq link','count>90','count>100'])
df_cera_ops=Final_df[col_ops]

df_cera_ops_dup=df_cera_ops.sort_values('count>90').drop_duplicates(subset='uniq link', keep='last')

table_ops = pd.pivot_table(df_cera_ops_dup, values='count>90', index=['uniq link'],  aggfunc=np.max)


#p= Final_df[['Circle','LT','uniq link','count>90','count>100']]
#p['Utilization'] = p['Utilization'].replace(np.nan,0)


#su = pd.pivot_table(p, index =['Circle','LT','uniq link'],values =['count>90','count>100'],aggfunc=max)
#su= su.reset_index()
#su = su.sort_values(by='count>90', ascending=False)

print (" Writing Start " )

writer = pd.ExcelWriter(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera Daily Utilization\OUTPUT\TXN_VIL_CERAGON_MW_Link_Utilization_Weekly_'+M+'.xlsx')
df_cera_ops_dup.to_excel(writer, sheet_name='Details',index=False)
Final_df.to_excel(writer, sheet_name='Backup',index=False)
#df.to_excel(writer, sheet_name='Total_Backup',index=False)

writer.close()

print (" Writing Done " )



print("Login Cobra")

ssh3=paramiko.SSHClient()
ssh3.set_missing_host_key_policy(paramiko.AutoAddPolicy())
try:
    ssh3.connect(hostname='10.115.1.57',username='Cobra',password='Cobra@123',port=22)
except:
    pass
try:
    ssh3.connect(hostname='10.19.62.229',username='Cobra',password='Cobra@123',port=22)
except:
    pass
sftp_client1=ssh3.open_sftp()

sftp_client1.chdir('/opt/MyLog/TX/MW_Link_Utilization_Daily')
sftp_client1.put(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera Daily Utilization\OUTPUT\TXN_VIL_CERAGON_MW_Link_Utilization_Weekly_'+M+'.xlsx', 'TXN_VIL_CERAGON_MW_Link_Utilization_Weekly_'+M+'.xlsx')


sftp_client1.close
ssh3.close

print("Upload done")








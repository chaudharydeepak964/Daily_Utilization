print ( "Welcome to Ceragon PM files Combined for UTIL Report " )

import pandas as pd
import numpy as np
import paramiko 
import datetime
from datetime import datetime, timedelta
import win32com.client as wincl
from openpyxl import load_workbook
import os
import glob
import warnings
import re
warnings.filterwarnings("ignore")

# Auto date


do=(datetime.now() - timedelta(1)).strftime('%d_%m_%Y')
di=(datetime.now() - timedelta(1)).strftime('%Y%m%d')
dm=(datetime.now() - timedelta(1)).strftime('%m-%y')
dn=(datetime.now() - timedelta(0)).strftime('%d-%m-%Y')
dnn=(datetime.now() - timedelta(1)).strftime('%d.%m.%Y')
dnnn=(datetime.now() - timedelta(1)).strftime('%Y-%m-%d')
DA = (datetime.now() - timedelta(0)).strftime('%d-%b-%y')



DA0= datetime.now() - timedelta(0)
DA1= datetime.now() - timedelta(1)
DA2= datetime.now() - timedelta(2)
DA3 = datetime.now() - timedelta(3)
DA4 = datetime.now() - timedelta(4)
DA5= datetime.now() - timedelta(5)
DA6= datetime.now() - timedelta(6)
DA7= datetime.now() - timedelta(7)


dnnn0= datetime.now() - timedelta(0)
dnnn1= datetime.now() - timedelta(1)
dnnn2= datetime.now() - timedelta(2)
dnnn3 = datetime.now() - timedelta(3)
dnnn4 = datetime.now() - timedelta(4)
dnnn5= datetime.now() - timedelta(5)
dnnn6= datetime.now() - timedelta(6)
dnnn7= datetime.now() - timedelta(7)

da7= datetime.now() - timedelta(1)
da6= datetime.now() - timedelta(2)
da5 = datetime.now() - timedelta(3)
da4 = datetime.now() - timedelta(4)
da3 = datetime.now() - timedelta(5)
da2 = datetime.now() - timedelta(6)
da1 = datetime.now() - timedelta(7)

do7= datetime.now() - timedelta(1)
do6= datetime.now() - timedelta(2)
do5 = datetime.now() - timedelta(3)
do4 = datetime.now() - timedelta(4)
do3 = datetime.now() - timedelta(5)
do2 = datetime.now() - timedelta(6)
do1 = datetime.now() - timedelta(7)


dn7= datetime.now() - timedelta(1)
dn6= datetime.now() - timedelta(2)
dn5 = datetime.now() - timedelta(3)
dn4 = datetime.now() - timedelta(4)
dn3 = datetime.now() - timedelta(5)
dn2 = datetime.now() - timedelta(6)
dn1 = datetime.now() - timedelta(7)


dnn7= datetime.now() - timedelta(1)
dnn6= datetime.now() - timedelta(2)
dnn5 = datetime.now() - timedelta(3)
dnn4 = datetime.now() - timedelta(4)
dnn3 = datetime.now() - timedelta(5)
dnn2 = datetime.now() - timedelta(6)
dnn1 = datetime.now() - timedelta(7)

DA7 = datetime.strftime(DA7, '%d-%b-%y')
DA6 = datetime.strftime(DA6, '%d-%b-%y')
DA5 = datetime.strftime(DA5, '%d-%b-%y')
DA4 = datetime.strftime(DA4, '%d-%b-%y')
DA3 = datetime.strftime(DA3, '%d-%b-%y')
DA2 = datetime.strftime(DA2, '%d-%b-%y')
DA1 = datetime.strftime(DA1, '%d-%b-%y')
DA0 = datetime.strftime(DA0, '%d-%b-%y')

da7 = datetime.strftime(da7, '%d_%m_%Y')
da6 = datetime.strftime(da6, '%d_%m_%Y')
da5 = datetime.strftime(da5, '%d_%m_%Y')
da4 = datetime.strftime(da4, '%d_%m_%Y')
da3 = datetime.strftime(da3, '%d_%m_%Y')
da2 = datetime.strftime(da2, '%d_%m_%Y')
da1 = datetime.strftime(da1, '%d_%m_%Y')

do7 = datetime.strftime(do7, '%Y%m%d')
do6 = datetime.strftime(do6, '%Y%m%d')
do5 = datetime.strftime(do5, '%Y%m%d')
do4 = datetime.strftime(do4, '%Y%m%d')
do3 = datetime.strftime(do3, '%Y%m%d')
do2 = datetime.strftime(do2, '%Y%m%d')
do1 = datetime.strftime(do1, '%Y%m%d')

dn7 = datetime.strftime(dn7, '%d-%m-%Y')
dn6 = datetime.strftime(dn6, '%d-%m-%Y')
dn5 = datetime.strftime(dn5, '%d-%m-%Y')
dn4 = datetime.strftime(dn4, '%d-%m-%Y')
dn3 = datetime.strftime(dn3, '%d-%m-%Y')
dn2 = datetime.strftime(dn2, '%d-%m-%Y')
dn1 = datetime.strftime(dn1, '%d-%m-%Y')

dnn7 = datetime.strftime(dnn7, '%d.%m.%Y')
dnn6 = datetime.strftime(dnn6, '%d.%m.%Y')
dnn5 = datetime.strftime(dnn5, '%d.%m.%Y')
dnn4 = datetime.strftime(dnn4, '%d.%m.%Y')
dnn3 = datetime.strftime(dnn3, '%d.%m.%Y')
dnn2 = datetime.strftime(dnn2, '%d.%m.%Y')
dnn1 = datetime.strftime(dnn1, '%d.%m.%Y')

dnnn0 = datetime.strftime(dnnn0, '%Y-%m-%d')
dnnn1 = datetime.strftime(dnnn1, '%Y-%m-%d')
dnnn2 = datetime.strftime(dnnn2, '%Y-%m-%d')
dnnn3 = datetime.strftime(dnnn3, '%Y-%m-%d')
dnnn4 = datetime.strftime(dnnn4, '%Y-%m-%d')
dnnn5 = datetime.strftime(dnnn5, '%Y-%m-%d')
dnnn6 = datetime.strftime(dnnn6, '%Y-%m-%d')

dm7=(datetime.now() - timedelta(1)).strftime('%d_%m_%Y')
dm6=(datetime.now() - timedelta(2)).strftime('%d_%m_%Y')
dm5=(datetime.now() - timedelta(3)).strftime('%d_%m_%Y')
dm4=(datetime.now() - timedelta(4)).strftime('%d_%m_%Y')
dm3=(datetime.now() - timedelta(5)).strftime('%d_%m_%Y')
dm2=(datetime.now() - timedelta(6)).strftime('%d_%m_%Y')
dm1=(datetime.now() - timedelta(7)).strftime('%d_%m_%Y')

da=(datetime.now() - timedelta(1)).strftime('%d_%m_%Y')
do=(datetime.now() - timedelta(1)).strftime('%Y%m%d')
dn=(datetime.now() - timedelta(1)).strftime('%d-%m-%Y')


print(dm7)
print(dm1)
print(da7)


print (" fILES rEADING sTART")



# KAR - IP-20

KAR=pd.read_csv(r'C:/Users/COR1736664/Desktop/Deepak/ALL CODE/Cera Daily Utilization/RAW/PM/K_Ethernet_Radio_Report_IP-20_24h_'+dm7+'.csv',usecols=['System Name','IP','Slot Number','Interface','Date','Peak Throughput'])

KAR['Server']='KAR'

print (' KAR Done ')


# ROB- IP -20

r7 = pd.read_csv(r'C:/Users/COR1736664/Desktop/Deepak/ALL CODE/Cera Daily Utilization/RAW/PM/R_Ethernet_Radio_Report_IP-20_24h_' + dm7 + '.csv', 
                 usecols=['System Name', 'IP', 'Slot Number', 'Interface', 'Date', 'Peak Throughput'],
                 encoding='ISO-8859-1')

##r7=pd.read_csv(r'C:/Users/COR1736664/Desktop/Deepak/ALL CODE/Cera Daily Utilization/RAW/PM/R_Ethernet_Radio_Report_IP-20_24h_'+dm7+'.csv',usecols=['System Name','IP','Slot Number','Interface','Date','Peak Throughput'])

# ROB- IP -10

r70=pd.read_csv(r'C:/Users/COR1736664/Desktop/Deepak/ALL CODE/Cera Daily Utilization/RAW/PM/R_Ethernet_Radio_Report_IP-10_24h_'+dm7+'.csv',usecols=['System Name','IP','Slot Number','Interface','Date','Peak Throughput'])

ROB=pd.concat([r7,r70])

ROB['Server']='ROB'

print (' ROB Done ')


# UPE- IP -20

u7=pd.read_csv(r'C:/Users/COR1736664/Desktop/Deepak/ALL CODE/Cera Daily Utilization/RAW/PM/U_Ethernet_Radio_Report_IP-20_24h_'+dm7+'.csv',usecols=['System Name','IP','Slot Number','Interface','Date','Peak Throughput'],encoding= 'unicode_escape')
#u7=pd.read_csv(r'C:/Users/COR1736664/Desktop/Deepak/ALL CODE/Cera Daily Utilization/RAW/PM/U_Ethernet_Radio_Report_IP-20_24h_'+dm7+'.csv',usecols=['System Name','Slot Number','Port Number','UAS'],error_bad_lines=False,encoding='latin1')

# UPE- IP -10

u70=pd.read_csv(r'C:/Users/COR1736664/Desktop/Deepak/ALL CODE/Cera Daily Utilization/RAW/PM/U_Ethernet_Radio_Report_IP-10_24h_'+dm7+'.csv',usecols=['System Name','IP','Slot Number','Interface','Date','Peak Throughput'],encoding= 'unicode_escape')

UPE=pd.concat([u7,u70])

UPE['Server']='UPE'

print (' UPE Done ')

print('All Files Read')


#Concat All Circle Dataframe
#Final_1=pd.concat([ASM,BIH,KAR,ROB,UPE])

dataframes = []
for name in ['KAR', 'ROB', 'UPE']:
    try:
        df = globals()[name]
        if isinstance(df, pd.DataFrame):
            dataframes.append(df)
    except KeyError:
        pass

Final_1 = pd.concat(dataframes, ignore_index=True)

print('Concat All Circle Dataframe Done')

##  ************** Extract number after "Slot" and "Radio #" *************

Final_1['Slot'] = Final_1['Slot Number'].str.extract(r'Slot (\d+)')
#Final_1['Port'] = Final_1['Port Number'].str.extract(r'Radio #(\d+)')
Final_1['Port']=Final_1['Interface'].str[-1]
Final_1['Port'] = Final_1['Port'].replace(['t'], '1')

##  ************** Create Short Name *************

Final_1['Short_Name']=Final_1['System Name']+','+Final_1['Slot']+','+Final_1['Port']

##  ************** Convert Peak Throughput *************


Final_1['Unit']=Final_1['Peak Throughput'].str[-4:]
Final_1['Unit'] = Final_1['Unit'].str.replace(' ','')

Final_1['Throughput'] = Final_1['Peak Throughput'].replace({'Mbps': '', 'Kbps': '', 'bps': ''}, regex=True)

# BPS
BPS=Final_1.loc[Final_1['Unit']=='bps']
BPS['Throughput']=pd.to_numeric(BPS['Throughput']) # change dtype from object to INT
BPS['Throughput']=BPS['Throughput']/1000000

# KBPS
KBPS=Final_1.loc[Final_1['Unit']=='Kbps']
KBPS['Throughput']=pd.to_numeric(KBPS['Throughput']) # change dtype from object to INT
KBPS['Throughput']=KBPS['Throughput']/1000

# MBPS
MBPS=Final_1.loc[Final_1['Unit']=='Mbps']

df=pd.concat([BPS,KBPS,MBPS])

##  ************** Remove unwanted space *************

df['Date'] = df['Date'].str.replace(' ','')


df = df.reindex(columns=['Short_Name','IP','System Name','Slot Number','Interface','Date','Peak Throughput','Unit','Throughput'])


##  ************** Create 7 New data frame *************


First_Day = df[df['Date'] == DA1]


print("* One Day Dataframe Created")


#Final_1 = Final_1.to_excel(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon PM File Availability\Cera_Combined.xlsx')
print("* Writing Start ")

writer = pd.ExcelWriter(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera Daily Utilization\OUTPUT\Final_PM_Combined_'+da7+'.xlsx')

df.to_excel(writer, sheet_name='Total',index=False)
First_Day.to_excel(writer, sheet_name= DA1,index=False)

writer.close()

print('Writing Done')

print("Report download done on local")

















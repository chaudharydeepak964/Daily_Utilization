print("CERAGON Daily Utilization REPORT STARTED")

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
da6=(datetime.now() - timedelta(2)).strftime('%Y%m%d')
da5=(datetime.now() - timedelta(3)).strftime('%Y%m%d')
da4=(datetime.now() - timedelta(4)).strftime('%Y%m%d')
da3=(datetime.now() - timedelta(5)).strftime('%Y%m%d')
da2=(datetime.now() - timedelta(6)).strftime('%Y%m%d')
da1=(datetime.now() - timedelta(7)).strftime('%Y%m%d')
print(da7)
print(da1)
dm7=(datetime.now() - timedelta(1)).strftime('%d_%m_%Y')
dm6=(datetime.now() - timedelta(2)).strftime('%d_%m_%Y')
dm5=(datetime.now() - timedelta(3)).strftime('%d_%m_%Y')
dm4=(datetime.now() - timedelta(4)).strftime('%d_%m_%Y')
dm3=(datetime.now() - timedelta(5)).strftime('%d_%m_%Y')
dm2=(datetime.now() - timedelta(6)).strftime('%d_%m_%Y')
dm1=(datetime.now() - timedelta(7)).strftime('%d_%m_%Y')
print(dm7)
print(dm1)

da=(datetime.now() - timedelta(1)).strftime('%d_%m_%Y')
do=(datetime.now() - timedelta(1)).strftime('%Y%m%d')
dn=(datetime.now() - timedelta(1)).strftime('%d-%m-%Y')

M=(datetime.now() - timedelta(1)).strftime('%d%m%Y')

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


rob = pd.read_csv("C:/Users/COR1736664/Desktop/Deepak/ALL CODE/Cera Daily Utilization/RAW/ROB_Full_Link_Report_"+da+".csv",skiprows=5,encoding= 'unicode_escape')
kar= pd.read_csv("C:/Users/COR1736664/Desktop/Deepak/ALL CODE/Cera Daily Utilization/RAW/KAR_Full_Link_Report_"+da+".csv",skiprows=5)
upe = pd.read_csv("C:/Users/COR1736664/Desktop/Deepak/ALL CODE/Cera Daily Utilization/RAW/UPE_Full_Link_Report_"+da+".csv",skiprows=5,encoding= 'unicode_escape')


print ( " Files Read " )



kar=kar[['Site A Name','Site A Physical Port','Site Z Name','Site Z Physical Port','Link Configuration','Site A IP','Site Z IP','Site A Radio','Site Z Radio','Site A Radio Script',
         'Site Z Radio Script','MRMC Script Profile','MRMC Script Maximum Profile','MRMC Script Minimum Profile','Site A Peak Throughput (Last 24 h) [Mb/s]',
         'Site Z Peak Throughput (Last 24 h) [Mb/s]','Site A Active E1/T1','Site Z Active E1/T1']]

rob=rob[['Site A Name','Site A Physical Port','Site Z Name','Site Z Physical Port','Link Configuration','Site A IP','Site Z IP','Site A Radio','Site Z Radio','Site A Radio Script',
         'Site Z Radio Script','MRMC Script Profile','MRMC Script Maximum Profile','MRMC Script Minimum Profile','Site A Peak Throughput (Last 24 h) [Mb/s]',
         'Site Z Peak Throughput (Last 24 h) [Mb/s]','Site A Active E1/T1','Site Z Active E1/T1']]

upe=upe[['Site A Name','Site A Physical Port','Site Z Name','Site Z Physical Port','Link Configuration','Site A IP','Site Z IP','Site A Radio','Site Z Radio','Site A Radio Script',
         'Site Z Radio Script','MRMC Script Profile','MRMC Script Maximum Profile','MRMC Script Minimum Profile','Site A Peak Throughput (Last 24 h) [Mb/s]',
         'Site Z Peak Throughput (Last 24 h) [Mb/s]','Site A Active E1/T1','Site Z Active E1/T1']]

print ( " Rename Column Name ")


upe.rename(columns={'Site Z Name':'Site B Name','Site Z Physical Port':'Site B Physical Port','Site Z IP':'Site B IP','Site Z Radio':'Site B Radio',
                    'Site Z Radio Script':'Site B Radio Script','Site A Peak Throughput (Last 24 h) [Mb/s]':'Site A Peak Throughput (Last 24 hours) [Mb/s]',
                    'Site Z Peak Throughput (Last 24 h) [Mb/s]':'Site B Peak Throughput (Last 24 hours) [Mb/s]','Site Z Active E1/T1':'Site B Active E1/T1'},inplace=True)

rob.rename(columns={'Site Z Name':'Site B Name','Site Z Physical Port':'Site B Physical Port','Site Z IP':'Site B IP','Site Z Radio':'Site B Radio',
                    'Site Z Radio Script':'Site B Radio Script','Site A Peak Throughput (Last 24 h) [Mb/s]':'Site A Peak Throughput (Last 24 hours) [Mb/s]',
                    'Site Z Peak Throughput (Last 24 h) [Mb/s]':'Site B Peak Throughput (Last 24 hours) [Mb/s]','Site Z Active E1/T1':'Site B Active E1/T1'},inplace=True)

kar.rename(columns={'Site Z Name':'Site B Name','Site Z Physical Port':'Site B Physical Port','Site Z IP':'Site B IP','Site Z Radio':'Site B Radio',
                    'Site Z Radio Script':'Site B Radio Script','Site A Peak Throughput (Last 24 h) [Mb/s]':'Site A Peak Throughput (Last 24 hours) [Mb/s]',
                    'Site Z Peak Throughput (Last 24 h) [Mb/s]':'Site B Peak Throughput (Last 24 hours) [Mb/s]','Site Z Active E1/T1':'Site B Active E1/T1'},inplace=True)




upe['Server']='UPE'
rob['Server']='ROB'
kar['Server']='KAR'


#df=pd.concat([asm,bih,odi,upe,rob,kar])

dataframes = []
for name in ['upe', 'rob','kar']:
    try:
        df = globals()[name]
        if isinstance(df, pd.DataFrame):
            dataframes.append(df)
    except KeyError:
        pass

df = pd.concat(dataframes, ignore_index=True)


# **** Update Link type ****

df['LT']='1+0'
df['LT'] = np.where(df['Link Configuration'].str.contains(r'Xpic|2\+0'), 'Xpic', '1+0')


# **** Finding Modulation and Channel Spaccing as per Radio Script ****

df['Site A Radio Script'] = df['Site A Radio Script'].fillna(df['Site B Radio Script'])

df['Site A Radio Script'] = df['Site A Radio Script'].astype(str)
df['QAM'] = df['Site A Radio Script'].apply(lambda x: re.search(r'(\d+QAM)', x).group(1) if re.search(r'(\d+QAM)', x) else None)

df['CHANNEL SPACING'] = df['Site A Radio Script'].apply(lambda x: re.search(r'(\d+(\.\d+)?)MHz', x).group(0) if re.search(r'(\d+(\.\d+)?)MHz', x) else None)


pattern_mapping = {
    '1414': '14MHz',
    '2828': '28MHz',
    '4040': '40MHz',
    '5656': '56MHz',
    '250250': '25MHz',
    '028028': '28MHz',
    '056056': '56MHz'
}

def map_channel_spacing(script):
    for pattern, spacing in pattern_mapping.items():
        if pattern in script:
            return spacing
    return None  # If no pattern matched

# Update missing values in 'CHANNEL SPACING' column
df['CHANNEL SPACING'] = df.apply(
    lambda row: map_channel_spacing(row['Site A Radio Script']) if pd.isna(row['CHANNEL SPACING']) else row['CHANNEL SPACING'],
    axis=1
)



# **** Finding Modulation as per MRMC Script Profile ****


def convert_to_numeric(value):
    try:
        return pd.to_numeric(value)
    except ValueError:
        return value

df['MRMC Script Profile'] = df['MRMC Script Profile'].apply(convert_to_numeric)
df['MRMC Script Maximum Profile'] = df['MRMC Script Maximum Profile'].apply(convert_to_numeric)
df['MRMC Script Minimum Profile'] = df['MRMC Script Minimum Profile'].apply(convert_to_numeric)


# Exclude 50CX and 50E
#dff = df.loc[df['Site A Radio']!='RFU-50CX']
dff = df.loc[~df['Site A Radio'].isin(['RFU-50CX', 'RFU-50E'])]

# FOR IP-10 and 20
MRMC_order = {0:'QPSK',1:'8QAM',2:'16QAM',3:'32QAM',4:'64QAM',5:'128QAM',6:'256QAM',
		7:'512QAM',8:'1024QAMLight',9:'1024QAM',10:'2048QAM',11:'2048QAM',12:'4096QAM'}

dff['Modulation1']=dff['MRMC Script Profile'].map(MRMC_order)
dff['Modulation2']=dff['MRMC Script Maximum Profile'].map(MRMC_order)
dff['Modulation3']=dff['MRMC Script Minimum Profile'].map(MRMC_order)

dff = dff.reset_index(drop=True)

## Find Final modulation on the basic of three columns

dff['Mod'] = dff['Modulation1'].fillna(dff['Modulation2'].combine_first(dff['Modulation3']))

dff['Modulation'] = dff['QAM'].fillna(dff['Mod'])

dff['Modulation'].fillna('QPSK',inplace=True)

## --------------------------------------------------------------------------------------------------------

# Only 50CX
df50=df.loc[df['Site A Radio']=='RFU-50CX']
df50['MRMC Script Profile'] = df50['MRMC Script Profile'].fillna(df50['MRMC Script Maximum Profile'])


# FOR IP-50
MRMC_order = {1:'4QAM',2:'8QAM',3:'16QAM',4:'32QAM',5:'64QAM',6:'128QAM',7:'256QAM',
		8:'512QAM',9:'1024QAMLight',10:'1024QAM',11:'2048QAM',12:'4096QAM',0:'2QAM'}


df50['Modulation1']=df50['MRMC Script Profile'].map(MRMC_order)
df50['Modulation2']=df50['MRMC Script Maximum Profile'].map(MRMC_order)
df50['Modulation3']=df50['MRMC Script Minimum Profile'].map(MRMC_order)

df50 = df50.reset_index(drop=True)

## Find Final modulation on the basic of three columns

df50['Mod'] = df50['Modulation1'].fillna(df50['Modulation2'].combine_first(df50['Modulation3']))

df50['Modulation'] = df50['QAM'].fillna(df50['Mod'])

df50['Modulation'].fillna('QPSK',inplace=True)


## --------------------------------------------------------------------------------------------------------

# Only 50CX
#dff50=df.loc[df['Site A Radio']=='RFU-50CX']
dff50=df.loc[df['Site A Radio'].isin(['RFU-50E'])]


# FOR IP-50
MRMC_order = {0:'BPSK9',1:'BPSK10',2:'BPSK',3:'QPSK',4:'8QAM',5:'16QAM',6:'32QAM',7:'64QAM',
		8:'128QAM',9:'256QAM',10:'512QAM'}


dff50['Modulation1']=dff50['MRMC Script Profile'].map(MRMC_order)
dff50['Modulation2']=dff50['MRMC Script Maximum Profile'].map(MRMC_order)
dff50['Modulation3']=dff50['MRMC Script Minimum Profile'].map(MRMC_order)

dff50 = dff50.reset_index(drop=True)

## Find Final modulation on the basic of three columns

dff50['Mod'] = dff50['Modulation1'].fillna(dff50['Modulation2'].combine_first(dff50['Modulation3']))

dff50['Modulation'] = dff50['QAM'].fillna(dff50['Mod'])

dff50['Modulation'].fillna('QPSK',inplace=True)



## --------------------------------------------------------------------------------------------------------

Comb = pd.concat([dff,df50,dff50])

df = Comb.copy()

df['Site A Name']=df['Site A Name'].str.strip()
df['Site B Name']=df['Site B Name'].str.strip()


# In[314]:


# unique link*******************************
df['uniq link']=np.where((df['Site A Name']<df['Site B Name']),(df['Site A Name']+'-'+df['Site B Name']),(df['Site B Name']+'-'+ df['Site A Name']))
df['uniq link']=df['uniq link'].str.strip()



# Finding Circle using UNIQUE LINK

df.loc[df['uniq link'].str.contains('IDDL|INDL',na=False),'Circle']='DEL'
df.loc[df['uniq link'].str.contains('IDUW|INUW',na=False),'Circle']='UPW'
df.loc[df['uniq link'].str.contains('IDOD|INOD',na=False),'Circle']='ODI'
df.loc[df['uniq link'].str.contains('IDKL|INKL',na=False),'Circle']='KEL'
df.loc[df['uniq link'].str.contains('IDAS|IDNE|INAS|INNE', na=False), 'Circle'] = 'ASM'
df.loc[df['uniq link'].str.contains('IDUE|INUE|AZMG',na=False),'Circle']='UPE'
df.loc[df['uniq link'].str.contains('IDKA|INKA|MYS0|MYS9',na=False),'Circle']='KAR'
df.loc[df['uniq link'].str.contains('IDWB|IINW|INEW|INWB',na=False),'Circle']='ROB'
df.loc[df['uniq link'].str.contains(' INB|BBSN|BCHN|BDAR|BLXM|BMRU|bnir|BPIA|BR10|BSAS|BTOD|BUGR|IDB0|IDBR|INBR|JBKU|KOLA|BN2083|BPPK',na=False),'Circle']='BIH'
df.loc[df['uniq link'].str.contains('ARJ0|Bhaw|CHK0|CNR0|DMP0|HWY1|IDJK|INJK|JMU0|jmu1|JMU2|NAG0|RAJ0|SRN0|SRR1|VIJ0',na=False),'Circle']='JNK'

dff=df.copy()


# In[324]:


dff[['slot A','PortA']] = dff['Site A Physical Port'].str.split("/",expand=True)
dff[['slot B','PortB']] = dff['Site B Physical Port'].str.split("/",expand=True)


# Fill blanks value with Port 1

dff['PortA'].fillna('Port 1',inplace=True)
dff['PortB'].fillna('Port 1',inplace=True)


# Replace Value in Column

dff['slot A'] = dff['slot A'].replace(['Slot 0'], 'Slot 1')
dff['slot B'] = dff['slot B'].replace(['Slot 0'], 'Slot 1')



dff['slot A']=dff['slot A'].str.strip().str[-1]
dff['slot B']=dff['slot B'].str.strip().str[-1]

dff['PortA']=dff['PortA'].str.strip().str[-1]
dff['PortB']=dff['PortB'].str.strip().str[-1]

dff['slot A'] = dff['slot A'].replace(['0'], '10')
dff['slot B'] = dff['slot B'].replace(['0'], '10')


print (" Creating Short Name")

#dff['Site A Short name'] = dff['Site A Name'].fillna('')+','+dff['slot A'].fillna('')+','+dff['PortA'].fillna('')

dff['Site A Short name'] = dff['Site A Name']+','+dff['slot A']+','+dff['PortA']
dff['Site B Short name'] = dff['Site B Name']+','+dff['slot B']+','+dff['PortB']

print ( " ALL DATAFRAME CREATED ")


dff['MAX_Throughput']=dff[['Site A Peak Throughput (Last 24 hours) [Mb/s]','Site B Peak Throughput (Last 24 hours) [Mb/s]']].max(axis=1)

dff['MAX_E1']=dff[['Site A Active E1/T1','Site B Active E1/T1']].max(axis=1)

dff.drop(columns='Site A Radio Script',inplace=True)
dff.drop(columns='Site B Radio Script',inplace=True)
dff.drop(columns='MRMC Script Profile',inplace=True)
dff.drop(columns='MRMC Script Maximum Profile',inplace=True)
dff.drop(columns='MRMC Script Minimum Profile',inplace=True)


print ( " Read Combined Pm File " )

pm1 = pd.read_excel(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera Daily Utilization\OUTPUT\Final_PM_Combined_'+dm7+'.xlsx',sheet_name=DA1,skiprows=0)
pm1.drop(columns=['IP','System Name','Slot Number','Interface','Date','Peak Throughput','Unit'],inplace=True)
pm1 = pm1.drop_duplicates(subset=['Short_Name']) # Using due to BIH 

# Merging Start ---------------------

pm1.rename(columns={'Short_Name':'Site A Short name','Throughput':'Throughput A'},inplace=True)
dff=pd.merge(dff,pm1,on='Site A Short name',how='left') # Merge 1


pm1.rename(columns={'Site A Short name':'Site B Short name','Throughput A':'Throughput B'},inplace=True)
dff=pd.merge(dff,pm1,on='Site B Short name',how='left') # Merge 2

# Merging Done  ---------------------

dff['MAX_TP']=dff[['Throughput A','Throughput B']].max(axis=1)

#Fill NaN (blank) value in MAX_TP column with MAX_Throughput column
dff['MAX_TP'] = dff['MAX_TP'].fillna(dff['MAX_Throughput'])


# ** Find Capacity **

#Capacity_order = {'2QAM':20,'4QAM':40,'8QAM':60,'16QAM':85,'32QAM':108,'64QAM':134,'128QAM':159,'256QAM':181,'512QAM':194,'1024QAM':215,'1024QAMLight':226,
#                  '2048QAM':240,'32QAM':108,'4096QAM':300}


# Exclude 50CX and 50E

Cdff = dff.loc[dff['Site A Radio']!='RFU-50E']
#Cdff = dff.loc[~dff['Site A Radio'].isin(['RFU-50CX', 'RFU-50E'])]

Capacity_order = {'QPSK':42,'2QAM':20,'4QAM':40,'8QAM':60,'16QAM':85,'32QAM':108,'64QAM':134,'128QAM':159,'256QAM':181,'512QAM':194,'1024QAMLight':215,'1024QAM':215,
                  '2048QAM':240,'4096QAM':300}
              	           
Cdff['Capacity']=Cdff['Modulation'].map(Capacity_order)

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
Cdff['Total_Capacity'] = Cdff.apply(calculate_capacity, axis=1)


## --------------------------------------------------------------------------------------------------------

# Only 50CX
Cdf50=dff.loc[dff['Site A Radio']=='RFU-50E']

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

Comb = pd.concat([Cdff,Cdf50])
dff = Comb.copy()

# Find Utilization
dff['MulE1s']=dff['MAX_E1'].multiply(other = 2.048, fill_value = 0)
dff['Tot_Tp']=round(dff['MAX_TP']+dff['MulE1s'],2)
dff['Utilization']=round(dff['Tot_Tp']/dff['Total_Capacity']*100,2)
dff[['Utilization']] =dff[['Utilization']].clip(upper=100)

dff = dff.sort_values(by='Utilization', ascending=False)

dff.drop(columns=['MAX_Throughput','Throughput A','Throughput B'],inplace=True)

dff = dff.reindex(columns=['Server','Circle','LT','Link Configuration','uniq link','Site A Name','Site B Name','Site A IP','Site B IP','Site A Physical Port',
                           'Site B Physical Port','Site A Radio','Site B Radio','Site A Short name','Site B Short name','CHANNEL SPACING','MAX_E1','Modulation','Capacity','Total_Capacity',
                           'MAX_TP','Tot_Tp','Utilization'])

# Remove blanks from Utilization
Final_df = dff[~dff['Utilization'].isna() & (dff['Utilization'] != '')]



print (" Pivot Start " )

p= Final_df[['Circle','LT','uniq link','Utilization']]

p['Utilization'] = p['Utilization'].replace(np.nan,0)


su = pd.pivot_table(p, index =['Circle','LT','uniq link'],values =['Utilization'],aggfunc=max)
su= su.reset_index()
su = su.sort_values(by='Utilization', ascending=False)

print (" Writing Start " )

writer = pd.ExcelWriter(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera Daily Utilization\OUTPUT\TXN_VIL_CERAGON_MW_Link_Utilization_Daily_'+M+'.xlsx')
su.to_excel(writer, sheet_name='Details',index=False)
Final_df.to_excel(writer, sheet_name='Backup',index=False)
dff.to_excel(writer, sheet_name='Total_Backup',index=False)

writer.close()

print (" Writing Done " )



print("Login Cobra")

ssh3=paramiko.SSHClient()
ssh3.set_missing_host_key_policy(paramiko.AutoAddPolicy())
try:
    ssh3.connect(hostname='10.10.10.10',username='admin',password='admin',port=22)
except:
    pass

try:
    ssh3.connect(hostname='11.11.11.11',username='admin',password='admin',port=22)
except:
    pass

sftp_client1=ssh3.open_sftp()

sftp_client1.chdir('/opt/MyLog/TX/MW_Link_Utilization_Daily')
sftp_client1.put(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera Daily Utilization\OUTPUT\TXN_VIL_CERAGON_MW_Link_Utilization_Daily_'+M+'.xlsx', 'TXN_VIL_CERAGON_MW_Link_Utilization_Daily_'+M+'.xlsx')


sftp_client1.close
ssh3.close

print("Upload done")



# Upload code BI PORTAL --- 


print(" ** Uploading Start on BI Portal ** ")

ssh3=paramiko.SSHClient()
ssh3.set_missing_host_key_policy(paramiko.AutoAddPolicy())
try:
    ssh3.connect(hostname='1.1.1.1',username='root',password='root',port=22)
except:
    pass

sftp_client1=ssh3.open_sftp()

sftp_client1.chdir('/home/snenrc/VIL_IDEA_REPORTS/TX_REPORT')
sftp_client1.put(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera Daily Utilization\OUTPUT\TXN_VIL_CERAGON_MW_Link_Utilization_Daily_'+M+'.xlsx', 'TXN_VIL_CERAGON_MW_Link_Utilization_Daily_'+M+'.xlsx')

sftp_client1.close
ssh3.close

print(" ** Congratulation Report uploaded on BI Portal ** ")



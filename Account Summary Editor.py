#!/usr/bin/env python
# coding: utf-8

# Transform and Combine Multiple Excel Files

# In[115]:


# Import dependencies
import pandas as pd
import os
from pathlib import Path
import win32 as win32


# In[116]:


# Directory Paths
raw_data_path = os.path.abspath(r'D:\MY FILES\Account Summary-HDFC')
final_data = os.path.abspath(r'D:\Files')
processed_path = os.path.abspath(r'D:\MY FILES\Account Summary-HDFC\Processed')

#Files of raw folder
files=os.listdir(raw_data_path)
files


# In[117]:


#Column Name to be
column_name = ['Date',
 'Narration',
 'Chq./Ref.No.',
 'Value Dt',
 'Withdrawal Amt.',
 'Deposit Amt.',
 'Closing Balance']
# column_name = data.iloc[19].to_list()


# In[118]:


# Adding try block to open existing file 
try:
    df= pd.read_excel(f'{final_data}/Account_Summary_HDFC.xlsx','Sheet1')
except:
    df=pd.DataFrame()


# In[119]:


# Remove the unwanted rows
for file in files:
    if file.endswith('.xls'):
        source_file=f'{raw_data_path}/{file}'

        # Reading the excel file into dataframe
        data = pd.read_excel(source_file, sheet_name='Sheet 1')

        # Renaming the Column Headings
        data = data.set_axis(column_name, axis=1)

        # Removes the first 21 rows and last 26 rows
        data= data.iloc[21:-26]

        #Appending the files into single data frame
        df= df.append(data)

        #Move the processed file into Processed folder
        os.replace(source_file,f'{processed_path}/{file}')

#Saving the file into Final folder
df.to_excel(f'{final_data}/Account_Summary_HDFC.xlsx', index=False)







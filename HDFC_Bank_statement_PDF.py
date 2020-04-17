
# coding: utf-8

# In[2]:


import pandas as pd
import tabula


# In[3]:


hdfc_statement='sample_pdf.pdf'


# In[4]:


hdfc=tabula.read_pdf(hdfc_statement,pages='all', pandas_options={'header': None})


# In[6]:


print(type(hdfc), len(hdfc))
type(hdfc[0])


# In[8]:


hf0 = hdfc[0]
print(type(hf0),len(hf0))


# In[9]:


hf0


# In[10]:


cnt =  len(hdfc)
total_hdfc =hdfc[0].append(hdfc[1:cnt])


# In[11]:


total_hdfc.info()


# In[12]:


total_hdfc.shape


# In[13]:


total_hdfc


# In[14]:


total_hdfc.reset_index(inplace = True,drop = True)


# In[15]:


total_hdfc


# In[16]:


total_hdfc.rename(columns={0: 'Date', 1: 'Narration', 2:'Chq', 3:'ValueDt', 4:'WithDrawalAmt', 5:'DepositAmt', 6:'ClosingBalance'}, inplace=True)


# In[17]:


total_hdfc


# In[18]:


total_hdfc[2:3]


# In[19]:


for index, row in total_hdfc.iterrows():
    #skip the first record
    if index==0:
        continue
    
    x=row.Date
    
    if pd.isnull(x):
        print('Index: ',index)
        
        # Append Previous record of Narration + current Narration
        Narration_msg=total_hdfc[index-1:index]['Narration']+row.Narration
        print(Narration_msg)
        
        # Update Previous Record
        total_hdfc[index-1:index]['Narration'] = Narration_msg


# In[20]:


total_hdfc


# In[21]:


# Drop the first record
total_hdfc=total_hdfc.drop(0)


# In[22]:


total_hdfc.head()


# In[23]:


# drop record if date filed is null
final_hdfc=total_hdfc.dropna(subset=['Date'])


# In[24]:


final_hdfc


# In[25]:


final_hdfc.reset_index(inplace = True,drop = True)


# In[26]:


final_hdfc


# In[27]:


final_hdfc.tail(10)


# In[28]:


final_hdfc.to_csv('final_hdfc.csv',index = False, header=True)


# In[95]:


df=pd.read_csv('final_hdfc1.csv')


# In[96]:


df.tail(10)


# In[97]:


df['DepositAmt']=df['DepositAmt'].str.replace(',', '').astype(float)


# In[98]:


df['WithDrawalAmt']=df['WithDrawalAmt'].str.replace(',', '').astype(float)


# In[99]:


df['ClosingBalance']=df['ClosingBalance'].str.replace(',', '').astype(float)


# In[100]:


df.head()


# ### Connect to database

# In[31]:


import sqlite3


# In[40]:


conn=sqlite3.connect('Bank_Statement.db')  #Database name is "Bank_Statement.db"


# In[41]:


c=conn.cursor()


# In[42]:


c.execute('''CREATE TABLE bank_stt (Date date, Narration varchar2, Chq varchar2, ValueDt date, 
          WithDrawalAmt float, DepositAmt float, ClosingBalance float)''')
conn.commit()


# In[43]:


df.to_sql('bank_stt', conn, if_exists='replace', index = False)


# In[45]:


c.execute('''  
SELECT * FROM bank_stt
          ''')


# In[46]:


for row in c.fetchall():
    print(row)


# ### Working With Excel

# In[111]:


import openpyxl


# In[51]:


# Give the location of the file
#path = "C:\\Users\\Itishree\\Downloads\\ICICI_Stmt.xlsx"


# In[112]:


#workbook = openpyxl.load_workbook(path)
workbook =  openpyxl.load_workbook('ICICI_Stmt.xlsx')


# In[115]:


worksheet = workbook.active


# In[105]:


Total_Credit=df['DepositAmt'].sum()
print(Total_Credit)


# In[117]:


worksheet['D3']=Total_Credit


# In[123]:


Balance=df.iloc[-1][-1]
print(Balance)


# In[124]:


worksheet['C3']=Balance


# In[144]:


#df['ClosingBalance'].sum() / df.index.size
Avg_Bal_6months=df['ClosingBalance'].mean()
print(Avg_Bal_6months)


# In[145]:


worksheet['B11']=Avg_Bal_6months


# In[146]:


workbook.save('ICICI_Stmt.xlsx')


# In[171]:


df.dtypes


# In[169]:


#df['Date']=df['Date'].astype('date')
df['Date']= pd.to_datetime(df['Date'])


# In[170]:


df.head()


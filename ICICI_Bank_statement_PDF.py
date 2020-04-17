
# coding: utf-8

# In[127]:


import pandas as pd
import tabula


# In[61]:


statement='ICICI_Statement_6months.pdf'


# In[62]:


bank_st=tabula.read_pdf(statement,pages='all', pandas_options={'header': None})


# In[63]:


print(type(bank_st), len(bank_st))
type(bank_st[0])


# In[64]:


bk0 = bank_st[0]
print(type(bk0),len(bk0))


# In[65]:


bk0.head(10)


# In[95]:


cnt =  len(bank_st)
total_state =bank_st[0].append(bank_st[1:cnt])


# In[97]:


total_state.info()


# In[98]:


total_state.shape


# In[99]:


total_state


# In[100]:


total_state.reset_index(inplace = True,drop = True)


# In[101]:


total_state


# In[102]:


total_state = total_state.drop([0], axis=1)


# In[103]:


total_state.head()


# In[104]:


total_state.rename(columns={1: 'Date', 4: 'Narration', 3:'Chq', 2:'ValueDt', 5:'WithDrawalAmt', 6:'DepositAmt', 7:'ClosingBalance'}, inplace=True)


# In[105]:


total_state.head()


# In[106]:


# Drop the first 4 records
total_state = total_state.drop(total_state.index[0:4])


# In[116]:


# drop record if Narration filed is null
total_state=total_state.dropna(subset=['Narration'])


# In[117]:


total_state.reset_index(inplace = True,drop = True)


# In[118]:


myIndex = 0
for index, row in total_state.iterrows():
    x = row.Date
    
    if pd.notnull(x):
        myIndex = index      
        print("reset",myIndex)
    else:
        print("Index -", index )   
        
        # Append Previous record of Narration + current Narration
        narMsg= total_state[myIndex:myIndex+1]['Narration'] + row.Narration
        
        # Update Previous Record
        total_state[myIndex:myIndex+1]['Narration'] = narMsg
        
        print(narMsg)


# In[114]:


total_state.head()


# In[119]:


# drop record if date filed is null
total_state=total_state.dropna(subset=['Date'])


# In[123]:


total_state.reset_index(inplace = True,drop = True)


# In[124]:


total_state


# In[131]:


total_state.to_csv('total_state_ICICI.csv',index = False, header=True)


# In[132]:


df=pd.read_csv('total_state_ICICI.csv')


# In[135]:


df


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


# In[136]:


Total_Credit=df['DepositAmt'].sum()
print(Total_Credit)


# In[117]:


worksheet['D3']=Total_Credit


# In[138]:


Balance=df.iloc[-1][-1]
print(Balance)


# In[124]:


worksheet['C3']=Balance


# In[137]:


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


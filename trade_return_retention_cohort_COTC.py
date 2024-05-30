#!/usr/bin/env python
# coding: utf-8

# In[1]:


# import
import pandas as pd
import numpy as np
import warnings
import duckdb
import seaborn as sns
import win32com.client
import time


# In[2]:


# preferences
warnings.filterwarnings("ignore")


# In[3]:


# damage data
damage_df=pd.read_excel(open('damage_report_2022.xlsx', 'rb'), sheet_name='Material Receving Status Report', header=0, index_col=None)


# In[4]:


# damage data slice
qry='''
select ClaimDate claim_date, DistributorCode distrib_code, SKUName material_desc, DamageReason reason, ReceivedAmount rec_amt
from damage_df; 
'''
damage_df=duckdb.query(qry).df()
damage_df


# In[5]:


# PH data
ph_df=pd.read_excel(open('01. PH2022-10th Jan 23.xlsx', 'rb'), sheet_name='Selling code_Jan 2023', header=0, index_col=None) 
qry='''
select distinct "Pack size desc." pack_desc, "Material Description" material_desc
from ph_df;
'''
ph_df=duckdb.query(qry).df()
ph_df


# In[7]:


# COTC/Premium/MD data
cotc_df=pd.read_excel(open('IOP 2023 - COTC, MD & Premium List - Final.xlsx', 'rb'), sheet_name='COTC, MD & Premium List', header=0, index_col=None) 
qry='''
select distinct upper(SKUs) pack_desc, "Category" cat, "Business Group" bg
from cotc_df
where
    "COTC in 2023?" ilike '%yes%'
    -- or "Premium in 2023?" ilike '%yes%'
    -- or "MD in 2023?" ilike '%yes%';
'''
cotc_df=duckdb.query(qry).df()
cotc_df['brand']=cotc_df['pack_desc'].apply(lambda x: x.split()[0])
cotc_df


# In[8]:


# COTC damage data
qry='''
select left(claim_date, 7) claim_month, distrib_code, reason, rec_amt
from 
    damage_df tbl1 
    inner join 
    ph_df tbl2 using(material_desc)
    inner join 
    cotc_df tbl3 using(pack_desc);
'''
ret_df=duckdb.query(qry).df()
ret_df


# In[9]:


# cohort
qry='''
with
    cohort as
    (select 
        tbl1.claim_month claim_month_from, 
        tbl2.claim_month claim_month_to, 
        datediff('month', concat(tbl1.claim_month, '-01')::date, concat(tbl2.claim_month, '-01')::date) month_diff, 
        count(distinct tbl1.distrib_code) distributors_ret
    from 
        (select claim_month, distrib_code, sum(rec_amt) rec_amt
        from ret_df
        where reason ilike '%expire%'
        group by 1, 2
        ) tbl1 

        left join 

        (select claim_month, distrib_code, sum(rec_amt) rec_amt
        from ret_df
        where reason ilike '%expire%'
        group by 1, 2
        ) tbl2 on(tbl1.distrib_code=tbl2.distrib_code and tbl1.claim_month<=tbl2.claim_month and tbl1.rec_amt<=tbl2.rec_amt)
    group by 1, 2, 3
    ) 
    
select *, distributors_ret*1.00/distributors_init distributors_ret_pct
from 
    cohort tbl1
    
    inner join 
    
    (select claim_month_from, distributors_ret distributors_init
    from cohort
    where month_diff=0
    ) tbl2 using(claim_month_from)
order by 1, 2; 
'''
cohort_df=duckdb.query(qry).df()
# cohort_df['distributors_ret_pct']=cohort_df['distributors_ret_pct'].map('{:,.2%}'.format)
cohort_df


# In[14]:


# style
cm=sns.color_palette("blend:white,blue", as_cmap=True)
cm.set_bad("white")


# In[15]:


# cohort (number)
cohort_df_piv=pd.pivot_table(cohort_df, index='claim_month_from', columns='month_diff', values='distributors_ret', aggfunc='sum')
cohort_df_piv.loc['avg.']=cohort_df_piv.mean(axis=0)
# show
cohort_df_piv.style.format("{:.0f}", na_rep='').background_gradient(cmap=cm, axis=None, low=0.1, high=1.0)


# In[16]:


# cohort (pct)
cohort_df_piv=pd.pivot_table(cohort_df, index='claim_month_from', columns='month_diff', values='distributors_ret_pct', aggfunc='sum')
cohort_df_piv.loc['avg.']=cohort_df_piv.mean(axis=0)
# show
cohort_df_piv.style.format("{:.2%}", na_rep='').background_gradient(cmap=cm, axis=None, low=0.1, high=1.0)


# In[17]:


# email
ol = win32com.client.Dispatch("outlook.application")
olmailitem = 0x0
newmail = ol.CreateItem(olmailitem)

# subject, recipients
newmail.Subject = 'COTC Damage Retention Cohort'
newmail.To = 'mehedi.asif@unilever.com'
newmail.CC = 'robius.sany@unilever.com'

# chart
html_tbl=(
    cohort_df_piv.style
    .format("{:.2%}", na_rep='')
    .set_properties(**{'width': '50px'}, **{'height': '20px'}, **{'font-size': '0.8em'}, **{'text-align': 'center'})
    .background_gradient(cmap=cm, axis=None, low=0.1, high=1.0)
    .render()
)

# body
newmail.HTMLbody = f'''
Hello Bhaiya,<br><br>
I have analyzed retention of distributors returning COTC SKUs throughout 2022. Alarmingly, >90% of ~195 distributors keep returning these on a monthly basis. The analysis is furnished below:<br><br>
''' + html_tbl + '''<br>
I am just curious if this behavior is normal, or whether this wastage should further be investigated. Please look at this critically and let me know your thoughts.<br><br>
Note that, the data was extracted at ''' + time.strftime('%d-%b-%y, %I:%M %p') + '''. This is an auto generated email using smtplib.<br><br>
Thanks,<br>
Shithi Maitra<br>
Asst. Manager, Cust. Service Excellence<br>
Unilever BD Ltd.<br>
'''

# display, send
# newmail.Display()
newmail.Send()


# In[ ]:





#!/usr/bin/env python
# coding: utf-8

# In[1]:


# import
import pandas as pd
import duckdb


# In[2]:


# depot stock by MDM
file_name=r'Daily_Stock_UBL_UCL_30 Mar 2023.XLSX'
depot_stock_df=pd.read_excel(open(file_name, 'rb'), sheet_name='Sheet1', header=0, index_col=None)

# distributor stock
file_name=r'Replenishment Repot_30 Mar 2023.xlsx'
norm_df=pd.read_excel(open(file_name, 'rb'), sheet_name='Replenishment UBL_UCL', header=0, index_col=None)

# volval
file_name=r'Mar TDP Region Vol-Val_ 01.03.23.xlsx'
volval_df=pd.read_excel(open(file_name, 'rb'), sheet_name='TDP', header=6, index_col=None)
cols = volval_df.columns.tolist()
for i in range(0, len(cols)): cols[i] = cols[i].replace('\n', ' ')
volval_df.columns = cols


# In[3]:


# shampoo tgt, norm, stk
qry='''
select 
    basepack, 
    norm_qty, norm_qty*val_per_cs/10000000 norm_val_crore, 
    depot_stock_qty, depot_stock_qty*val_per_cs/10000000 depot_stock_val_crore,
    customer_stock_qty, customer_stock_qty*val_per_cs/10000000 customer_stock_val_crore, 
    mar_tgt_crore
from 
    (select "Base pack" basepack, sum("Stock in transit"+"Available for orders") depot_stock_qty
    from depot_stock_df
    where 
        "Storage Loc."='BF01'
        and "Base pack" ilike '%shampoo%'
        and "Base pack" not ilike '%clinic%'
    group by 1
    ) tbl1 
    
    inner join 
    
    (select 
        Basepack basepack, 
        sum("Norm qty") norm_qty, 
        sum("Stock on hand"+"In transit"+"Open order") customer_stock_qty, 
        avg("Price / UOM") val_per_cs
    from norm_df
    where 
        Basepack ilike '%shampoo%'
        and Basepack not ilike '%clinic%'
    group by 1 
    ) tbl2 using(basepack)
    
    inner join

    (select SKU basepack, sum("Mar Val") mar_tgt_crore
    from volval_df
    group by 1
    ) tbl3 using(basepack)
'''
res_df=duckdb.query(qry).df()
res_df


# In[4]:


# excel
writer=pd.ExcelWriter(r"shampoo_norm_stk_tgt.xlsx", engine='xlsxwriter')
res_df.to_excel(writer, sheet_name="Sheet1", startcol=0, startrow=0, index=False)
writer.save()


# In[ ]:





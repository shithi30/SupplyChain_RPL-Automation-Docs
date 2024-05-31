#!/usr/bin/env python
# coding: utf-8

# In[1]:


# import
import pandas as pd
import duckdb


# In[2]:


# read RPL
rpl_df = pd.read_excel(open("Replenishment Repot_10 Sep 2023.xlsx", "rb"), sheet_name="Replenishment UBL_UCL", header=0, index_col=None)
rpl_df = rpl_df[['Town', 'Basepack', 'Proposed qty', 'Norm qty', 'Stock on hand', 'In transit', 'Open order']]
rpl_df.columns = ['town', 'basepack', 'proposed_qty', 'norm_qty', 'stock_on_hand', 'stock_transit', 'open_order']

# read classifications
ubl_cls_df = pd.read_excel(open("UBL Classification Sep.xlsx", "rb"), sheet_name="Town SKU Classification", header=0, index_col=None)
ucl_cls_df = pd.read_excel(open("UCL Classification Sep.xlsx", "rb"), sheet_name="Town SKU Classification", header=0, index_col=None)
cls_df = duckdb.query('''select * from ubl_cls_df union select * from ucl_cls_df''').df()
cls_df.columns = ['town', 'basepack_code', 'basepack', 'sale_val', 'contrib', 'cls']


# In[3]:


# town - portfolio (sku/abcd/overall) - sku_count - norm_qty - proposed_qty - stock_on_hand - value_index - qmix
qry = '''
with 
    tbl as 
    (select * 
    from 
        rpl_df tbl1 
        inner join 
        (select town, basepack, cls
        from cls_df 
        ) tbl2 using(town, basepack)
    where town in('NAOGAON', 'DHUPCHACHIYA', 'KASHINATHPUR', 'RANGPUR', 'DINAJPUR', 'SAIDPUR')
    )
    
-- basepack
select 
    town, basepack portfolio, 
    count(basepack) sku_count, sum(norm_qty) norm_qty, sum(proposed_qty) proposed_qty, sum(stock_on_hand) stock_on_hand, 
    sum(stock_on_hand)*1.00/sum(norm_qty) value_index, 1-sum(proposed_qty)*1.00/sum(norm_qty) qmix
from tbl
group by 1, 2

-- overall
union all 
select 
    town, 'overall' portfolio, 
    count(basepack) sku_count, sum(norm_qty) norm_qty, sum(proposed_qty) proposed_qty, sum(stock_on_hand) stock_on_hand, 
    sum(stock_on_hand)*1.00/sum(norm_qty) value_index, 1-sum(proposed_qty)*1.00/sum(norm_qty) qmix
from tbl
group by 1, 2

-- cls
union all 
select 
    town, cls portfolio, 
    count(basepack) sku_count, sum(norm_qty) norm_qty, sum(proposed_qty) proposed_qty, sum(stock_on_hand) stock_on_hand, 
    sum(stock_on_hand)*1.00/sum(norm_qty) value_index, 1-sum(proposed_qty)*1.00/sum(norm_qty) qmix
from tbl
group by 1, 2
'''
res_df = duckdb.query(qry).df()
display(res_df)

# summary
res_df_piv = duckdb.query('''select town, portfolio, value_index, qmix from res_df where portfolio in('A', 'B', 'C', 'D', 'overall')''').df()
res_df_piv = res_df_piv.pivot(index="town", columns="portfolio")
display(res_df_piv)


# In[5]:


# store
with pd.ExcelWriter("overstock_impact.xlsx") as writer:
    res_df.to_excel(writer, sheet_name="Full Data", index=False)
    res_df_piv.to_excel(writer, sheet_name="Summary", index=True)


# In[ ]:





# In[ ]:





# In[ ]:





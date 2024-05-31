#!/usr/bin/env python
# coding: utf-8

# In[1]:


# import
import pandas as pd
import numpy as np
import warnings
import duckdb
from glob import glob
import os
# preferences
warnings.filterwarnings("ignore")


# In[2]:


# SCCF data
path='Secondary_CCFOT_Report_'
sccf_df=pd.DataFrame()
for file in glob(path+'*.xlsx'):
    file_name=os.path.basename(file).split('.')[0]
    df=pd.read_excel(open(file, 'rb'), sheet_name='Sheet1', header=0, index_col=None)
    sccf_df=sccf_df.append(df)
    print("Read data from: "+file_name)


# In[3]:


# CP data
cp_df=pd.read_excel(open("02. PH2023-15th Jan 23.xlsx", 'rb'), sheet_name='Selling code_Jan 2023', header=0, index_col=None)
qry='''
select distinct material
from cp_df
where "Material Group"='F070'
'''
cp_df=duckdb.query(qry).df()


# In[4]:


# stock data
path='Secondary_Stock_Trend_Report_'
stock_dfs=[]
for file in glob(path+'*.xlsx'):
    file_name=os.path.basename(file).split('.')[0]
    df=pd.read_excel(open(file, 'rb'), sheet_name='Sheet1', header=0, index_col=None)
    stock_dfs.append(df)
    print("Read data from: "+file_name)


# In[12]:


# dates
qry='''
select report_date
from 
    (select distinct "Orig.Req.Delivery Date" report_date, strptime("Orig.Req.Delivery Date", '%d.%m.%Y')
    from sccf_df
    order by 2
    ) tbl1 
'''
date_df=duckdb.query(qry).df()
dates=date_df['report_date'].values.tolist()
l=len(dates)
dates


# In[71]:


sccf_df=sccf_df.replace('', '0')


# In[74]:


# accumulate
acc_label_df=pd.DataFrame()
for i in range(2, l): 
    # delivery 
    report_date=dates[i]
    # print(report_date)
    qry='''
    select 
        "Orig.Req.Delivery Date" report_date, 
        "Local Sales Region 1 (S.Sales)" region,
        "Local Sales Region 4" town, 
        "Material" material, 
        "Unnamed: 10" material_desc,
        "Pack Size" pack_size,
        "Final\nOrder Qty." final_order_qty,
        "Invoiced\nQuantity" invoiced_qty,
        "Final\nOrder Qty."-"Invoiced\nQuantity" sccf_loss_qty,
        "Case Fill %" sccf
    from sccf_df
    where "Orig.Req.Delivery Date"='''+"'"+report_date+"'"+'''
    '''
    delivery_df=duckdb.query(qry).df()

    # order 
    report_date=dates[i-1]
    # print(report_date)
    qry='''
    select 
        "Orig.Req.Delivery Date" report_date, 
        "Local Sales Region 1 (S.Sales)" region,
        "Local Sales Region 4" town, 
        "Material" material, 
        "Unnamed: 10" material_desc,
        "Pack Size" pack_size,
        "Final\nOrder Qty." final_order_qty,
        "Invoiced\nQuantity" invoiced_qty,
        "Final\nOrder Qty."-"Invoiced\nQuantity" sccf_loss_qty,
        "Case Fill %" sccf
    from sccf_df
    where "Orig.Req.Delivery Date"='''+"'"+report_date+"'"+'''
    '''
    order_df=duckdb.query(qry).df()
    
    # closing stock
    report_date=dates[i-2]
    # print(report_date)
    for df in stock_dfs:
        cols=df.columns
        if report_date in cols:
            stock_df=df
            break
    qry='''
    select "Unnamed: 3", "Unnamed: 4", "Unnamed: 7", "Unnamed: 8", "Unnamed: 9"
    from stock_df
    '''
    stock_common_cols_df=duckdb.query(qry).df()
    clstock_df=stock_df.filter(regex=report_date)
    clstock_df=stock_common_cols_df.join(clstock_df)
    clstock_df=clstock_df[1:] 
    clstock_df.columns=['region', 'town', 'material', 'material_desc', 'pack_size', 'stock_vol_cs']  

    # CP idef
    qry='''
    select 
        *, 
        case 
            -- when tbl2.material in (69575333, 69575326, 69575319, 69575314, 69572616, 69572607, 69572625, 69560905, 69560914, 69560935, 68932785) then 'running CP'
            when tbl2.material is not null then 'old CP' 
            else 'non CP' 
        end cp_status
    from 
        delivery_df tbl1 
        left join 
        cp_df tbl2 using(material)
    '''
    delivery_cp_idef_df=duckdb.query(qry).df()

    qry='''
    select 
        *, 
        case 
            -- when tbl2.material in (69575333, 69575326, 69575319, 69575314, 69572616, 69572607, 69572625, 69560905, 69560914, 69560935, 68932785) then 'running CP'
            when tbl2.material is not null then 'old CP' 
            else 'non CP' 
        end cp_status
    from 
        order_df tbl1 
        left join 
        cp_df tbl2 using(material)
    '''
    order_cp_idef_df=duckdb.query(qry).df()

    qry='''
    select 
        *, 
        case 
            -- when tbl2.material in (69575333, 69575326, 69575319, 69575314, 69572616, 69572607, 69572625, 69560905, 69560914, 69560935, 68932785) then 'running CP'
            when tbl2.material is not null then 'old CP' 
            else 'non CP' 
        end cp_status
    from 
        clstock_df tbl1 
        left join 
        cp_df tbl2 using(material)
    '''
    clstock_cp_idef_df=duckdb.query(qry).df()

    # labels
    qry='''
    select *
    from 
        (select *
        from delivery_cp_idef_df 
        where sccf<100
        ) tbl1

        left join 

        (select cp_status, region, town, pack_size, sum(stock_vol_cs) clstock_vol_cs
        from clstock_cp_idef_df
        group by 1, 2, 3, 4
        ) tbl2 using(cp_status, region, town, pack_size)

        left join 

        (select cp_status, region, town, pack_size, sum(invoiced_qty) given_order_qty
        from order_cp_idef_df
        group by 1, 2, 3, 4
        ) tbl3 using(cp_status, region, town, pack_size)
    '''
    label_df=duckdb.query(qry).df()

    c=label_df.select_dtypes(np.number).columns
    label_df[c]=label_df[c].fillna(0)

    qry='''
    select 
        *, 
        clstock_vol_cs-given_order_qty revised_stock, 
        case 
            when final_order_qty>revised_stock and sccf_loss_qty<=(final_order_qty-revised_stock) then sccf_loss_qty
            when final_order_qty>revised_stock and sccf_loss_qty>(final_order_qty-revised_stock) then final_order_qty-revised_stock
        end stock_loss_qty,
        case 
            when final_order_qty>revised_stock and sccf_loss_qty<=(final_order_qty-revised_stock) then 'full stock loss'
            when final_order_qty>revised_stock and sccf_loss_qty>(final_order_qty-revised_stock) then 'partial stock loss'
            when final_order_qty<=revised_stock then 'service loss'
            else 'unidentified loss'
        end loss_label
    from label_df
    '''
    label_df=duckdb.query(qry).df()

    acc_label_df=acc_label_df.append(label_df)
    print("Rows found for "+dates[i]+": "+str(label_df.shape[0]))
    
print() 
print("Total rows found: "+str(acc_label_df.shape[0]))


# In[93]:


# scorecard data
scorecard_df=pd.read_excel(open("sccf_2022.xlsx", 'rb'), sheet_name='Sheet1', header=0, index_col=None)
scorecard_df


# In[97]:


# primary DR data
primdr_df=pd.read_excel(open("Pri DR  Jan-Jun 2022 BP.xlsx", 'rb'), sheet_name='Sheet1', header=0, index_col=None)
primdr_df.columns


# In[138]:


# analysis 01: Surf
qry='''
select *, sccf+(stock_loss_qty_pct*(1-sccf)) potential_sccf
from 
    (select 
        town, 
        sum(sccf_loss_qty) sccf_loss_qty, 
        sum(stock_loss_qty) stock_loss_qty,
        sum(stock_loss_qty)/sum(sccf_loss_qty) stock_loss_qty_pct
    from acc_label_df
    where 
        pack_size='SURF NM STD POWDER EXCEL 500G'
        and stock_loss_qty>0
    group by 1
    ) tbl1
    
    left join
    
    (select 
        "Local Sales Region 4" town, 
        sum("Invoiced\nQuantity")/sum("Final\nOrder Qty.") sccf
    from sccf_df
    where "Pack Size"='SURF NM STD POWDER EXCEL 500G'
    group by 1
    ) tbl2 using(town)
'''
res_df=duckdb.query(qry).df()
res_df


# In[136]:


# analysis 02: pack size
qry='''
select *, sccf+(stock_loss_qty_pct*(1-sccf)) potential_sccf
from 
    (select 
        pack_size, 
        sum(sccf_loss_qty) sccf_loss_qty, 
        sum(stock_loss_qty) stock_loss_qty,
        sum(stock_loss_qty)/sum(sccf_loss_qty) stock_loss_qty_pct
    from 
        (select 
            pack_size, 
            final_order_qty, 
            case 
                when stock_loss_qty is null then 0
                else stock_loss_qty
            end stock_loss_qty,
            sccf_loss_qty
        from acc_label_df
        where sccf_loss_qty>0
        ) tbl1 
    group by 1
    ) tbl1 

    inner join 

    (select 
        "Pack Size (PH)" pack_size, 
        sum("Dispatched Qty.")/sum("Final Customer Expected Order Qty.") prim_dr
    from primdr_df
    group by 1
    ) tbl2 using(pack_size)
    
    inner join

    (select 
        "Pack Size" pack_size, 
        sum("Invoiced\nQuantity")/sum("Final\nOrder Qty.") sccf
    from sccf_df
    group by 1
    ) tbl3 using(pack_size)
order by 2 desc, 4 desc
limit 50
'''
res_df=duckdb.query(qry).df()
res_df


# In[103]:


# analysis 03: month
qry='''
select *, sccf_card+(stock_loss_qty_pct*(1-sccf_card)) potential_sccf
from 
    (select 
        right(report_date, 7) report_month, 
        sum(sccf_loss_qty) sccf_loss_qty, 
        sum(stock_loss_qty) stock_loss_qty,
        sum(stock_loss_qty)/sum(sccf_loss_qty) stock_loss_qty_pct
    from acc_label_df
    group by 1
    order by 1
    ) tbl1 

    inner join 

    (select month report_month, SCCF_card sccf_card
    from scorecard_df
    ) tbl2 using(report_month)
'''
res_df=duckdb.query(qry).df()
res_df


# In[76]:


# excel
qry='''
select *
from acc_label_df
'''
label_df_exl=duckdb.query(qry).df()
label_df_exl.to_excel('labeled_sccf_loss_h1_22.xlsx', engine='xlsxwriter', index=False)


# In[ ]:





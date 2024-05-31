#!/usr/bin/env python
# coding: utf-8

# import
import pandas as pd
import duckdb

# depot stock by MDM
file_name = r'Daily_Stock_UBL_UCL_19 Mar 2023.XLSX'
depot_stock_df = pd.read_excel(open(file_name, 'rb'), sheet_name='Sheet1', header=0, index_col=None)

# norm
file_name=r'Replenishment Repot_19 Mar 2023.xlsx'
norm_df=pd.read_excel(open(file_name, 'rb'), sheet_name='Replenishment UBL_UCL', header=0, index_col=None)

# volval
file_name = r'Mar TDP Region Vol-Val_ 01.03.23.xlsx'
volval_df = pd.read_excel(open(file_name, 'rb'), sheet_name='TDP', header=6, index_col=None)
cols = volval_df.columns.tolist()
for i in range(0, len(cols)): cols[i] = cols[i].replace('\n', ' ')
volval_df.columns = cols

# plant/depot
file_name = r'plant_mapping.xlsx'
plantmap_df = pd.read_excel(open(file_name, 'rb'), sheet_name='Sheet1', header=0, index_col=None)

# daily
file_name = r'Monthly_Order_&_Allocation_Report_19 Mar 2023.xlsx'
daily_df = pd.read_excel(open(file_name, 'rb'), sheet_name='Summary BP', header=0, index_col=None)
daily_df = daily_df[['Basepack Description', 'All_information', daily_df.columns.to_list()[-18]]]
daily_df.columns = ['Basepack', 'all_info', 'last_day_stats']
qry = '''
select 
    Basepack,
    sum(case when all_info='Secondary Achievement' then last_day_stats else null end) "Secondary Achievement",
    sum(case when all_info='Business Contribution' then last_day_stats else null end) "Business Contribution"
from daily_df 
where all_info in('Secondary Achievement', 'Business Contribution')
group by 1
'''
daily_df = duckdb.query(qry).df()

# excel
writer = pd.ExcelWriter(r"0_but_live.xlsx", engine='xlsxwriter')

# depot stock 0
qry = '''
select * 
from 
    (select Plant depot_code, "Base pack" Basepack, sum("Stock in transit"+"Available for orders") depot_stock
    from depot_stock_df
    where "Storage Loc."='BF01'
    group by 1, 2
    having sum("Stock in transit"+"Available for orders")=0
    ) tbl0
    
    inner join 
    
    (select distinct Basepack
    from norm_df
    ) tbl1 using(Basepack)
    
    inner join 
    
    (select Plant depot_code, Plant_Name depot_name
    from plantmap_df
    ) tbl2 using(depot_code)
'''
res_df0 = duckdb.query(qry).df()
display(res_df0)
res_df0.to_excel(writer, sheet_name="depot_stock_0", startcol=0, startrow=0, index=False)

# depot stock 0, but live
qry = '''
select * 
from 
    (select "Customer Code", "Customer name", "Basepack Code", Basepack, "Norm qty", "Town", "Depot Name" depot_name, "Stock on hand"+"In transit"+"Open order" customer_stock
    from norm_df
    ) tbl1 
    
    inner join 
    
    (select Plant depot_code, Plant_Name depot_name
    from plantmap_df
    ) tbl2 using(depot_name)
    
    inner join 

    res_df0 tbl3 using(depot_code, Basepack)
'''
res_df = duckdb.query(qry).df()
display(res_df)
res_df.to_excel(writer, sheet_name="depot_stock_0_live", startcol=0, startrow=0, index=False)

# customer stock 0, but live
qry = '''
select * 
from 
    (select "Customer Code", "Customer name", "Basepack Code", Basepack, "Norm qty", "Town", "Depot Name" depot_name, "Stock on hand"+"In transit"+"Open order" customer_stock
    from norm_df
    where "Stock on hand"+"In transit"+"Open order"=0
    ) tbl1 
    
    inner join 
    
    (select Plant depot_code, Plant_Name depot_name
    from plantmap_df
    ) tbl2 using(depot_name)
    
    inner join 
    
    (select Plant depot_code, "Base pack" Basepack, sum("Stock in transit"+"Available for orders") depot_stock
    from depot_stock_df
    where "Storage Loc."='BF01'
    group by 1, 2
    ) tbl3 using(depot_code, Basepack)
'''
res_df2 = duckdb.query(qry).df()
display(res_df2)
res_df2.to_excel(writer, sheet_name="customer_stock_0_live", startcol=0, startrow=0, index=False)

# depot and customer stock 0, but live
qry = '''
select * 
from 
    (select "Customer Code", "Customer name", "Basepack Code", Basepack, "Norm qty", "Town", "Depot Name" depot_name, "Stock on hand"+"In transit"+"Open order" customer_stock
    from norm_df
    where "Stock on hand"+"In transit"+"Open order"=0
    ) tbl1 
    
    inner join 
    
    (select Plant depot_code, Plant_Name depot_name
    from plantmap_df
    ) tbl2 using(depot_name)
    
    inner join 
    
    (select Plant depot_code, "Base pack" Basepack, sum("Stock in transit"+"Available for orders") depot_stock
    from depot_stock_df
    where "Storage Loc."='BF01'
    group by 1, 2
    having sum("Stock in transit"+"Available for orders")=0
    ) tbl3 using(depot_code, Basepack)
    
    left join 
    
    daily_df tbl4 using(Basepack)
where 
    "Secondary Achievement"<1
    and "Business Contribution" is not null
order by 12 desc, 11 asc
'''
res_df3 = duckdb.query(qry).df()
display(res_df3)
res_df3.to_excel(writer, sheet_name="depot+customer_stock_0_live", startcol=0, startrow=0, index=False)

# volval primary 0, secondary > 0, but live
qry = '''
select *
from 
    (select "Customer Code", "Customer name", "Basepack Code", Basepack, "Norm qty", "Town", "Depot Name" depot_name, "Stock on hand"+"In transit"+"Open order" customer_stock
    from norm_df
    ) tbl1 
    
    inner join 
    
    (select SKU Basepack, Mar volval_vol_prim, "Secondary Total (vol)" volval_vol_sec, "Remarks for CSE"
    from volval_df
    where "Secondary Total (vol)"=0 and Mar>0
    ) tbl2 using(Basepack)
'''
res_df4 = duckdb.query(qry).df()
display(res_df4)
res_df4.to_excel(writer, sheet_name="primary_0_sec_nonzero_live", startcol=0, startrow=0, index=False)

# analysis
qry = '''
select 
    count(Basepack) basepacks,
    sum("Business Contribution") "Business Contribution", 
    avg("Secondary Achievement") "Secondary Achievement"
from 
    (select distinct Basepack, "Secondary Achievement", "Business Contribution"
    from res_df3
    ) tbl1
'''
anls_df = duckdb.query(qry).df()
display(anls_df)

# save
writer.save()

#!/usr/bin/env python
# coding: utf-8

# import
import os
from glob import glob
import pandas as pd
import duckdb

# read RPL
rpl_df = pd.read_excel(open("Replenishment Repot_20 Sep 2023.xlsx", "rb"), sheet_name="Replenishment UBL_UCL", header=0, index_col=None)
rpl_df = rpl_df[['Town', 'Basepack', 'Stock on hand', 'Category description', 'Classification']]
rpl_df.columns = ['town', 'basepack', 'stock_on_hand', 'category', 'classification']
rpl_df = rpl_df.applymap(lambda s: s.upper() if type(s)==str else s)

# read SCCF
path = os.getcwd()
sccf_df = pd.DataFrame()
for file in glob(path + "\*SEC CCF*.xlsx"):
    file_name = os.path.basename(file)
    print("Reading data from: " + file_name)
    df = pd.read_excel(open(file, "rb"), sheet_name="Sheet1", header=1, index_col=None)
    df = df[['Local Sales Region 4', 'Pack Size', 'CS', 'CS.1']]
    df.columns = ['town', 'basepack', 'ord_qty', 'inv_qty']
    sccf_df = sccf_df.append(df)
sccf_df = sccf_df.applymap(lambda s: s.upper() if type(s)==str else s)

# read stock-allocation
stkalloc_df = pd.read_excel(open("Town x SKU  Allocation September'23.xlsx", "rb"), sheet_name="Town x SKU x Case x TGT ", header=2, index_col=None)
stkalloc_df = stkalloc_df[['SKU NAME', 'BUSINESS GROUP', 'CATEGORY', 'TOWN x SKU TGT - TP Cr.']]
stkalloc_df.columns = ['basepack', 'BG', 'category', 'tgt_cr' ]
stkalloc_df = stkalloc_df.applymap(lambda s: s.upper() if type(s)==str else s)

# read national classification
ubl_cls_df = pd.read_excel(open("UBL Classification Sep.xlsx", "rb"), sheet_name="National Classification", header=2, index_col=None)
ucl_cls_df = pd.read_excel(open("UCL Classification Sep.xlsx", "rb"), sheet_name="National Classification", header=2, index_col=None)
ntnl_cls_df = duckdb.query('''select Basepack, Classification from ubl_cls_df union select Basepack, Classification from ucl_cls_df''').df()
ntnl_cls_df.columns = ['basepack', 'classification']
ntnl_cls_df = ntnl_cls_df.applymap(lambda s: s.upper() if type(s)==str else s)

# read town classification
ubl_cls_df = pd.read_excel(open("UBL Classification Sep.xlsx", "rb"), sheet_name="Town SKU Classification", header=0, index_col=None)
ucl_cls_df = pd.read_excel(open("UCL Classification Sep.xlsx", "rb"), sheet_name="Town SKU Classification", header=0, index_col=None)
town_cls_df = duckdb.query('''select * from ubl_cls_df union select * from ucl_cls_df''').df()
town_cls_df.columns = ['town', 'basepack_code', 'basepack', 'sale_val', 'contrib', 'classification']
town_cls_df = town_cls_df.applymap(lambda s: s.upper() if type(s)==str else s)

# read national RR
rr_df = pd.read_excel(open("Working - September - Raw-L6M Sales History & RR Hana File_UBL & UCL.xlsx", "rb"), sheet_name="Sheet3", header=2, index_col=None)
rr_df = rr_df[['Row Labels', 'Sum of L3M Daily RR', 'Sum of L6M Daily RR']]
rr_df.columns = ['basepack', 'RR_3_months', 'RR_6_months']
rr_df = rr_df.applymap(lambda s: s.upper() if type(s)==str else s)

# town RR
town_rr_df = pd.read_excel(open("Working - September - Raw-L6M Sales History & RR Hana File_UBL & UCL.xlsx", "rb"), sheet_name="Sheet1 (2)", header=2, index_col=None)
town_rr_df = town_rr_df[['Local Sales Region 4(m.d.)', 'Pack Size(m.d.)', 'L6M Daily RR', 'L3M Daily RR']]
town_rr_df.columns = ['town', 'basepack', 'RR_3_months', 'RR_6_months']
qry = '''
select town, basepack, sum(RR_3_months) RR_3_months, sum(RR_6_months) RR_6_months
from town_rr_df 
group by 1, 2
'''
town_rr_df = duckdb.query(qry).df()
town_rr_df = town_rr_df.applymap(lambda s: s.upper() if type(s)==str else s)
display(town_rr_df)

# RR this month
days = len(glob(path + "\*SEC CCF*.xlsx"))-2
qry = '''
select town, basepack, sum(ord_qty)*1.00/''' + str(days) + ''' RR_this_month
from sccf_df 
group by 1, 2
'''
rr_thmon_df = duckdb.query(qry).df()
rr_thmon_df = rr_thmon_df.applymap(lambda s: s.upper() if type(s)==str else s)
display(rr_thmon_df)

# BC
qry = '''
select BG, category, basepack, sum(tgt_cr)*1.00/(select sum(tgt_cr) from stkalloc_df) basepack_bc 
from stkalloc_df
group by 1, 2, 3
'''
bc_df = duckdb.query(qry).df()
bc_df = bc_df.applymap(lambda s: s.upper() if type(s)==str else s)
display(bc_df)

# national data
qry = '''
select BG, category, classification, basepack_bc, basepack, attr, val
from
    (select basepack, 'stock_on_hand' attr, sum(stock_on_hand) val
    from rpl_df 
    group by 1, 2
    
    union all
    
    select basepack, 'DOH_3_months' attr, stock_on_hand*1.00/RR_3_months val
    from 
        (select basepack, sum(stock_on_hand) stock_on_hand
        from rpl_df 
        group by 1
        ) tbl1
        inner join 
        (select basepack, RR_3_months
        from rr_df
        ) tbl2 using(basepack)
        
    union all
    
    select basepack, 'DOH_6_months' attr, stock_on_hand*1.00/RR_6_months val
    from 
        (select basepack, sum(stock_on_hand) stock_on_hand
        from rpl_df 
        group by 1
        ) tbl1
        inner join 
        (select basepack, RR_6_months
        from rr_df
        ) tbl2 using(basepack)
        
    union all
    
    select basepack, 'DOH_this_month' attr, stock_on_hand*1.00/RR_this_month val
    from 
        (select basepack, sum(stock_on_hand) stock_on_hand
        from rpl_df 
        group by 1
        ) tbl1
        inner join 
        (select basepack, sum(RR_this_month) RR_this_month
        from rr_thmon_df
        group by 1
        ) tbl2 using(basepack)
    
    union all
    
    select distinct basepack, 'DOH_target' attr, 18 val
    from rpl_df
    ) tbl1
    
    left join
    
    (select BG, category, basepack, basepack_bc from bc_df) tbl2 using(basepack)
    
    left join 
    
    (select basepack, classification from ntnl_cls_df) tbl3 using(basepack)
'''
ntnl_df = duckdb.query(qry).df() #.fillna('')

# pivot
piv_df = pd.pivot_table(
    ntnl_df, 
    values='val', 
    index=['BG', 'category', 'classification', 'basepack_bc', 'basepack'], 
    columns=['attr'], 
    aggfunc='sum'
    ).reset_index()
display(piv_df)

# town data
qry = '''
select BG, category, town, classification, basepack_bc, basepack, attr, val
from
    (select town, basepack, 'stock_on_hand' attr, sum(stock_on_hand) val
    from rpl_df 
    group by 1, 2, 3
    
    union all
    
    select town, basepack, 'DOH_3_months' attr, stock_on_hand*1.00/RR_3_months val
    from 
        (select town, basepack, sum(stock_on_hand) stock_on_hand
        from rpl_df 
        group by 1, 2
        ) tbl1
        inner join 
        (select town, basepack, RR_3_months
        from town_rr_df
        ) tbl2 using(town, basepack)
        
    union all
    
    select town, basepack, 'DOH_6_months' attr, stock_on_hand*1.00/RR_6_months val
    from 
        (select town, basepack, sum(stock_on_hand) stock_on_hand
        from rpl_df 
        group by 1, 2
        ) tbl1
        inner join 
        (select town, basepack, RR_6_months
        from town_rr_df
        ) tbl2 using(town, basepack)
        
    union all
    
    select town, basepack, 'DOH_this_month' attr, stock_on_hand*1.00/RR_this_month val
    from 
        (select town, basepack, sum(stock_on_hand) stock_on_hand
        from rpl_df 
        group by 1, 2
        ) tbl1
        inner join 
        (select town, basepack, RR_this_month
        from rr_thmon_df
        ) tbl2 using(town, basepack)
    
    union all
    
    select town, basepack, 'DOH_target' attr, 18 val
    from rpl_df
    ) tbl1
    
    left join
    
    (select BG, category, basepack, basepack_bc from bc_df) tbl2 using(basepack)
    
    left join 
    
    (select town, basepack, classification from town_cls_df) tbl3 using(town, basepack)
'''
town_df = duckdb.query(qry).df() #.fillna('')

# pivot
piv2_df = pd.pivot_table(
    town_df, 
    values='val', 
    index=['BG', 'category', 'classification', 'basepack_bc', 'basepack', 'town'], 
    columns=['attr'], 
    aggfunc='sum'
    ).reset_index()
display(piv2_df)

# store
with pd.ExcelWriter("C:/Users/Shithi.Maitra/Downloads/operational_file.xlsx") as writer:
    piv_df.to_excel(writer, sheet_name="National", index=False)
    piv2_df.to_excel(writer, sheet_name="Town", index=False)

# # sanity check
# qry = '''
# select town, basepack, count(*) entry
# from rpl_df
# group by 1, 2
# order by 3 desc
# '''
# df = duckdb.query(qry).df()
# display(df)


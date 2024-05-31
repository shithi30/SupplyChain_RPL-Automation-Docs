#!/usr/bin/env python
# coding: utf-8

# import
import pandas as pd
import duckdb

# lifting plan
file = "C:/Users/Shithi.Maitra/Downloads/UBL_Lifting_plan_10 Nov 2023.xlsx"
sheet_name = "Lifting Plan"
lp_df = pd.read_excel(open(file, "rb"), sheet_name=sheet_name, header=0, index_col=None)
lp_df.columns = ['region', 'area', 'cust_code', 'town', 'plan', 'manual_plan', 'full_manual', 'inv_date', 'opd']
display(lp_df)

# confirmed plan
file = "C:/Users/Shithi.Maitra/Downloads/All Town confirm order Full_09 Nov 2023.xlsx"
sheet_name = "All Town confirm order"
cf_df = pd.read_excel(open(file, "rb"), sheet_name=sheet_name, header=0, index_col=None)
cf_df = cf_df[['Town', 'Customer confirm value']]
cf_df.columns = ['town', 'confirmed_val']
cf_df = duckdb.query('''select town, sum(confirmed_val) confirmed_val from cf_df group by 1''').df()
display(cf_df)

# read RPL
file = "C:/Users/Shithi.Maitra/Unilever Codes/Ad Hoc/2by2 Matrices/RPL Inputs/2023-11-09_Replenishment Repot_09 Nov 2023.xlsx"
df = pd.read_excel(open(file, "rb"), sheet_name="Replenishment UBL_UCL", header=0, index_col=None)
df = df[['Date', 'Town', 'Basepack', 'Stock on hand', 'Customer Inventory Status', 'Classification', 'Norm qty', 'Perday sales qty', 'Defaulter list', 'PPO value']]
df.columns = ['rpl_date', 'town', 'basepack', 'stock', 'inv_status', 'cls', 'norm_qty', 'daily_sales_qty', 'if_default', 'ppo_val']
df = duckdb.query('''select strptime(rpl_date, '%d %b %Y') rpl_date, upper(town) town, upper(basepack) basepack, stock, inv_status, cls, norm_qty, daily_sales_qty, stock*1.00/daily_sales_qty doh, ppo_val, if_default from df''').df()
rpl_df = df

# winter RPL
qry = '''
select rpl_date, town, basepack, cls, doh, norm_qty, stock, norm_qty-stock stock_to_norm_diff, stock*1.00/norm_qty stock_to_norm_pct
from rpl_df
where cls='Winter'
'''
winter_df = duckdb.query(qry).df()
display(winter_df)

# winter BCs
file = "C:/Users/Shithi.Maitra/Unilever Codes/Ad Hoc/2by2 Matrices/RPL Inputs/" + "Town x SKU  Allocation November'23.xlsx"
sheet_name = "Town x SKU x Case x TGT "
df = pd.read_excel(open(file, "rb"), sheet_name=sheet_name, header=2, index_col=None)
tgt_df = df[['TOWN NAME', 'SKU NAME', 'TOWN x SKU TGT - TP Cr.']]
tgt_df.columns = ['town', 'basepack', 'tgt_cr']
tgt_df = duckdb.query('''select upper(town) town, upper(basepack) basepack, tgt_cr from tgt_df where upper(basepack) in(select basepack from winter_df)''').df()
tgt_df = duckdb.query('''select *, tgt_cr*1.00/(select sum(tgt_cr) from tgt_df) bc from tgt_df''').df()
display(tgt_df)

# winter lines (stock < norm)
qry = '''
select * 
from 
    (select *
    from winter_df
    where stock_to_norm_diff>0
    ) tbl1 
    inner join 
    tgt_df tbl2 using(town, basepack)
'''
anls_df = duckdb.query(qry).df()
display(anls_df)

# towns eligible to be pushed to
qry = '''
select *, case when plan > confirmed_val then 'yes' else 'no' end if_push_eligible, plan-confirmed_val push_val
from 
    (select 
        town, 
        sum(stock_to_norm_diff) stock_to_norm_diff,
        sum(bc) winter_bc_of_low_stock
    from anls_df
    group by 1
    ) tbl1 
    
    inner join 
    
    (select town, plan
    from lp_df
    ) tbl2 using(town)
    
    inner join 
    
    (select town, confirmed_val*1.00/10000000 confirmed_val
    from cf_df
    ) tbl3 using(town)
order by stock_to_norm_diff desc
'''
step_df = duckdb.query(qry).df()
display(step_df)

# winter lines (stock < 10 DOH)
qry = '''
select * 
from 
    (select *
    from winter_df
    where doh < 10
    ) tbl1 
    inner join 
    tgt_df tbl2 using(town, basepack)
'''
anls2_df = duckdb.query(qry).df()
display(anls2_df)

# towns eligible to be pushed to
qry = '''
select *, case when plan > confirmed_val then 'yes' else 'no' end if_push_eligible, plan-confirmed_val push_val
from 
    (select 
        town, 
        sum(10-doh) doh_lag_from_10,
        sum(bc) winter_bc_of_low_stock
    from anls2_df
    group by 1
    ) tbl1 
    
    inner join 
    
    (select town, plan
    from lp_df
    ) tbl2 using(town)
    
    inner join 
    
    (select town, confirmed_val*1.00/10000000 confirmed_val
    from cf_df
    ) tbl3 using(town)
order by doh_lag_from_10 desc
'''
step2_df = duckdb.query(qry).df()
display(step2_df)

# store
with pd.ExcelWriter("C:/Users/Shithi.Maitra/Downloads/winter_low_stock.xlsx") as writer:
        anls_df.to_excel(writer, sheet_name="Winter Lines - Stock vs Norm", index=False)
        step_df.to_excel(writer, sheet_name="Action Towns - Stock vs Norm", index=False)
        anls2_df.to_excel(writer, sheet_name="Winter Lines - Stock vs 10 DOH", index=False)
        step2_df.to_excel(writer, sheet_name="Action Towns - Stock vs 10 DOH", index=False)



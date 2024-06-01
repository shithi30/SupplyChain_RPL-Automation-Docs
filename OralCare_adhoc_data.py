# Plz share the info below
# Oral Care SKU 
# 1. National stock - CS and MT 
# 2. Avg. order last 5 days: CS 
# 3. Monthly remaining volume

# import
import pandas as pd
import duckdb

# read MTD sales
file_name = r'Monthly_Order_&_Allocation_Report_BP_22 Mar 2023.xlsx'
daily_df = pd.read_excel(open(file_name, 'rb'), sheet_name='Summary BP', header=0, index_col=None)
daily_df.head(5)

# desired stats
qry = '''
select *, secondary_target_cs-primary_cs monthly_remaining_vol_cs
from 
    (select 
        "Basepack Description" basepack, 
        "2023-03-22" biz_contribution
    from daily_df
    where 
        All_information='Business Contribution'
        and ("Basepack Description" ilike 'close%' or "Basepack Description" ilike 'pepsodent%')
    order by 2 desc
    limit 10
    ) tbl1 
    
    inner join 
    
    (select 
        "Basepack Description" basepack, 
        "2023-03-22" national_stock_cs
    from daily_df
    where All_information='National Stock in CS'
    ) tbl2 using(basepack)
    
    inner join 
    
    (select 
        "Basepack Description" basepack, 
        ("2023-03-17"+"2023-03-18"+"2023-03-19"+"2023-03-20"+"2023-03-21")/5 last_5_day_avg_orders_cs
    from daily_df
    where All_information='Order Qnt in CS'
    ) tbl3 using(basepack)
    
    inner join 
    
    (select 
        "Basepack Description" basepack, 
        "2023-03-22" secondary_target_cs
    from daily_df
    where All_information='Secondary Target in CS'
    ) tbl4 using(basepack)
    
    inner join 
    
    (select 
        "Basepack Description" basepack, 
        "2023-03-01"+"2023-03-02"+"2023-03-03"+"2023-03-04"+"2023-03-05"+
        "2023-03-06"+"2023-03-07"+"2023-03-08"+"2023-03-09"+"2023-03-10"+
        "2023-03-11"+"2023-03-12"+"2023-03-13"+"2023-03-14"+"2023-03-15"+
        "2023-03-16"+"2023-03-17"+"2023-03-18"+"2023-03-19"+"2023-03-20"+
        "2023-03-21"+"2023-03-22" primary_cs
    from daily_df
    where All_information='Primary In CS'
    ) tbl5 using(basepack)
where basepack not ilike '%powder%'
order by 2 desc
'''
res_df = duckdb.query(qry).df()
display(res_df)


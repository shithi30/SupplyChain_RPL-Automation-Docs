#!/usr/bin/env python
# coding: utf-8

# import
import pandas as pd
import warnings
import duckdb

# preferences
warnings.filterwarnings("ignore")

# PH data
ph_df = pd.read_excel(open("03. PH2023-26th Jan 23.xlsx", 'rb'), sheet_name="Selling code_Jan 2023", header=0, index_col=None)

# active data 
active_df = pd.read_excel(open("rptActiveBasepack_14022023.xlsx", 'rb'), sheet_name="Sheet1", header=0, index_col=None)
qry='''
select distinct "Basepack Code" active_bp_code, "Material Code" active_material_code
from active_df
'''
active_df = duckdb.query(qry).df()
active_df.head(5)

# jupyter data sheet-01
jupyter_df1 = pd.read_excel(open("Feb'23 Jupiter communication - 13.02.23 (Revised).xlsx", 'rb'), sheet_name="February'23 Jupiter Activity", header=0, index_col=None)

# jupyter + bp_code data
qry = '''
select tbl1.*, tbl2.pack_size_code
from 
    jupyter_df1 tbl1 
    left join
    (select distinct "Pack size" pack_size_code, "Pack size desc." pack_size_desc
    from ph_df
    ) tbl2 on(tbl1."Base pack "=tbl2.pack_size_desc)
'''
jupyter_bpcode_df1 = duckdb.query(qry).df()

# priority 1
qry = '''
select tbl1.* 
from 
    (select "February Activity Code", "Base pack ", "Description", "Subs for activity code", pack_size_code
    from jupyter_bpcode_df1
    where "Subs for activity code"=1
    ) tbl1 
    
    left join 
    
    active_df tbl2 on(tbl1.pack_size_code=tbl2.active_bp_code and tbl1."February Activity Code"=tbl2.active_material_code)
where tbl2.active_bp_code is null
'''
res_df1 = duckdb.query(qry).df()
print(res_df1)

# jupyter data sheet-02
jupyter_df2 = pd.read_excel(open("Feb'23 Jupiter communication - 13.02.23 (Revised).xlsx", 'rb'), sheet_name="January'23 Slip Out", header=0, index_col=None)

# jupyter + bp_code data
qry= '''
select tbl1.*, tbl2.pack_size_code
from 
    jupyter_df2 tbl1 
    left join
    (select distinct "Pack size" pack_size_code, "Pack size desc." pack_size_desc
    from ph_df
    ) tbl2 on(tbl1."Base pack"=tbl2.pack_size_desc)
'''
jupyter_bpcode_df2 = duckdb.query(qry).df()
print(jupyter_bpcode_df2.shape[0])

# priority 1
qry = '''
select tbl1.* 
from
    (select 
        "Base pack", 
        "January Activity Code (Remove Substitution)", "Description", "January Post activity code   (Add Substitution)",
        "Post activity code substitution", 
        pack_size_code
    from jupyter_bpcode_df2
    where "Post activity code substitution"=1
    ) tbl1 

    left join 

    active_df tbl2 on(tbl1.pack_size_code=tbl2.active_bp_code and tbl1."January Post activity code   (Add Substitution)"=tbl2.active_material_code)
where tbl2.active_bp_code is null
'''
res_df2 = duckdb.query(qry).df()
print(res_df2)

# alloc data (trimmed)
alloc_df = pd.read_excel(open("AllocationDetails_Version 12 to 14 Feb 2023_UBL.xlsx", 'rb'), sheet_name="Allocation Details", header=2, index_col=None)
qry = '''
select distinct
    "Basepack Code" alloc_bp_code, "Basepack Description" alloc_bp_desc, 
    "Material Code" alloc_material_code, "Material Description" alloc_material_desc
from alloc_df
where 
    "Basepack Code" not in 
    (select pack_size_code
    from res_df1
    union 
    select pack_size_code
    from res_df2
    )
    and "Allocated Qty">0
'''
alloc_trim_df = duckdb.query(qry).df()
print(alloc_trim_df)

# price data
price_df = pd.read_excel(open("RadGridExport (26).xls", 'rb'), sheet_name="RadGridExport", header=0, index_col=None)
qry = '''
select *, rank() over(partition by price_bp_code order by price desc) seq
from 
    (select distinct "Variant Code" price_bp_code, Code price_material_code, Name price_material_desc, "Retail Price" price
    from price_df
    ) tbl1 
'''
price_df = duckdb.query(qry).df()
print(price_df)

# if active price < alloc price
qry = '''
select *
from 
    (select 
        alloc_bp_code::text alloc_bp_code, 
        alloc_bp_desc::text alloc_bp_desc,
        alloc_material_code::int alloc_material_code,
        alloc_material_desc::text alloc_material_desc
    from alloc_trim_df 
    ) tbl1 
    
    inner join 
    
    (select price_material_code alloc_material_code, price alloc_price
    from price_df
    ) tbl2 using(alloc_material_code)
    
    left join 
    
    (select active_bp_code alloc_bp_code, active_material_code
    from active_df
    ) tbl3 using(alloc_bp_code)
    
    left join
    
    (select price_material_code active_material_code, price_material_desc active_material_desc, price active_price
    from price_df
    ) tbl4 using(active_material_code)
where alloc_price>active_price
'''
active_greater_alloc_df = duckdb.query(qry).df()
print(active_greater_alloc_df)

# cases to activate
qry='''
select
    alloc_bp_code bp_code, 
    alloc_bp_desc bp_desc, 
    alloc_material_code material_code, 
    alloc_material_desc material_desc, 
    active_material_code,
    active_material_desc, 
    'allocation price > active price' remarks
from active_greater_alloc_df

union all

select 
    pack_size_code bp_code,
    "Base pack " bp_desc,
    "February Activity Code" material_code,
    "Description" material_desc,
    '' active_material_code,
    '' active_material_desc, 
    'jupiter sheet-01' remarks
from res_df1

union all

select
    pack_size_code bp_code,
    "Base pack" bp_desc,
    "January Post activity code   (Add Substitution)" material_code,
    "Description" material_desc,
    '' active_material_code,
    '' active_material_desc, 
    'jupiter sheet-02' remarks
from res_df2
'''
res_df = duckdb.query(qry).df()
print(res_df)

# to Excel
res_df.to_excel('activate_cases.xlsx', engine='xlsxwriter', index=False)


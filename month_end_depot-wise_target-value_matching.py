#!/usr/bin/env python
# coding: utf-8

# import
import pandas as pd
import numpy as np
import duckdb
from pathlib import Path
import openpyxl
from openpyxl import load_workbook, Workbook
import xlsxwriter
import win32com.client
from datetime import datetime, timedelta

# read lifting plan (LP)
file = "Invoicing Plan For 24th June'23 - UBL.xlsx"
sheet_name = "Sheet1"
df = pd.read_excel(open(file, "rb"), sheet_name=sheet_name, header=0, index_col=None)
df.columns = ['depot', 'town', 'plan1']
qry = '''
select depot, town, (case when plan1!='\n' then plan1 else 0 end)::numeric plan
from df
where town not like '%TOTAL%'
'''
lp_df = duckdb.query(qry).df()
display(lp_df)

# read allocation
file = "Allocation Report (2).xlsx"
sheet_name = "Summary"
df = pd.read_excel(open(file, "rb"), sheet_name=sheet_name, header=3, index_col=None)
qry = '''
select Town town, sum("Allocated Value") alloc
from df
group by 1
'''
alloc_df = duckdb.query(qry).df()
display(alloc_df)

# gap = plan - allocation
qry = '''
select *, gap*1.00/sum(gap) over(partition by depot) total_gap_pct
from 
    (select 
        depot, town, plan,
        case when alloc is null then 0 else alloc end alloc,
        plan-(case when alloc is null then 0 else alloc end)/10000000 gap
    from 
        lp_df tbl1 
        left join 
        alloc_df tbl2 using(town)
    ) tbl1
'''
gap_df = duckdb.query(qry).df()
display(gap_df)

# read C-class SKUs
file = "C Class SKU Classification.xlsx"
sheet_name = "Sheet1"
df = pd.read_excel(open(file, "rb"), sheet_name=sheet_name, header=3, index_col=None)
c_df = df[['Depot', 'Basepack']]
c_df.columns = ['depot1', 'basepack']
qry = '''
select distinct
    basepack, 
    case when depot1=='CTG' then 'Nur Jahan-CTG' else depot1 end depot
from c_df
'''
c_df = duckdb.query(qry).df()
display(c_df)

# read stock
file = "Daily_Stock_UBL_UCL_23 Jun 2023 (3).xlsx"
sheet_name = "Sheet1"
df = pd.read_excel(open(file, "rb"), sheet_name=sheet_name, header=0, index_col=None)
qry = '''
select * 
from 
    (select 
        case 
            when "Plant"='B102' then 'Bogra'
            when "Plant"='B104' then 'Barisal'
            when "Plant"='B105' then 'Khulna'
            when "Plant"='B111' then 'Dhaka'
            when "Plant"='B902' then 'Nur Jahan-CTG'
            else null 
        end depot,
        Basepack basepack, 
        max("Available for orders") stock
    from df
    where "Storage Loc."='BF01'
    group by 1, 2
    ) tbl1 

    inner join 

    c_df tbl2 using(depot, basepack)
'''
stock_df = duckdb.query(qry).df()
display(stock_df)

# match target
qry = '''
select *, stock*total_gap_pct to_fulfil, floor(stock*total_gap_pct) to_fulfil_rounded
from 
    gap_df tbl1 
    inner join 
    stock_df tbl2 using(depot)
'''
match_df = duckdb.query(qry).df()
display(match_df)

# result
match_df.to_excel("output.xlsx", index = False)

# # multiplicity check
# qry = '''
# select depot, basepack, count(*) instances
# from stock_df
# group by 1,2
# order by 3 desc
# '''
# df = duckdb.query(qry).df()
# display(df)




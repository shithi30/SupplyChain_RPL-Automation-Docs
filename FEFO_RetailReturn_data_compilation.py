#!/usr/bin/env python
# coding: utf-8

# import
import os
from glob import glob
import pandas as pd
import duckdb

# preference
duckdb.default_connection.execute("set global pandas_analyze_sample=100000")

# input dir
path = input("Enter the path containing the month's Excel files:\n")
path = (path + "\\").replace("\\", "/") # path = "C:/Users/Shithi.Maitra/Downloads/Project FEFO/Project FEFO/DAMAGE MIS 22-23/DAMAGE MIS 22-23/4. April 2022 Damage MIS/"
print("\nNOTE: All warnings are safe to ignore.\n")

# accumulators
visit_df = pd.DataFrame()
error_files = []
row_count = 0

# read PH
print("Reading PH data...\n")
file = "C:/Users/Shithi.Maitra/Downloads/1. PH2024 - 17th Jan 2024-UBL & UCL.xlsx"
ph_df = pd.read_excel(open(file, "rb"), sheet_name="Selling code Jan 2024", header=0, index_col=None)

# read outlets 
print("Reading outlet data...\n")
file = "C:/Users/Shithi.Maitra/Downloads/Project FEFO/Project FEFO/Latest Outlet List_11_04_23_Zia Bhai.xlsb"
outlet_df = pd.read_excel(open(file, "rb"), sheet_name="National_OSDS_Outlet_Data_237", header=0, index_col=None)

# read PSNS SKUs
print("Reading PS/NS data...\n")
file = "C:/Users/Shithi.Maitra/Downloads/Project FEFO/Project FEFO/PS NS SKUs 2023.xlsx"
ps_df = pd.read_excel(open(file, "rb"), sheet_name="PS", header=1, index_col=None)
ns_df = pd.read_excel(open(file, "rb"), sheet_name="NS", header=1, index_col=None)
qry = '''
select "SKIN CARE SKUs" sku, 'PS' psns_label from ps_df 
union all
select SKU, 'NS' psns_label from ns_df 
'''
psns_df = duckdb.query(qry).df()

# create FEFO
print("Creating FEFO data...\n")
fefo_df = pd.DataFrame()
fefo_df['sku'] = [
    'CLEAR MALE SHAMPOO CSM 330ML',
    'CLEAR SHAMPOO ANTI HAIR FALL 350ML',
    'SUNSILK SHAMPOO BLACK 375ML',
    'CLEAR SHAMPOO COMPLETE ACTVE CARE 350ML',
    'CLEAR SHAMPOO COMPLETE ACTVE CARE 180ML',

    'DOVE SHAMPOO HAIR FALL RESCUE 340ML',
    'VASELINE LOTION HLTHY WHTE 200ML',
    'SUNSILK SHAMPOO PERFECT STRAIGHT 375ML',
    'DOVE SHAMPOO INTENSIVE REPAIR 340ML',
    'CLEAR MALE SHAMPOO CSM 180ML',

    'SUNSILK SHAMPOO BLACK 180ML',
    'PEPSODENT TOOTH POWDER 100G',
    'GLOW & LOVELY FCL MSTRSR MV CRM 9G',
    'SUNSILK SHAMPOO HFS 375ML',
    'PEPSODENT TOOTH POWDER 50G',

    'CLEAR MALE SHAMPOO CSM 5ML',
    'DOVE SHAMPOO IRP 6ML',
    'GLOW & LOVELY FC WASH FM INSTA GLOW 100G',
    'GLOW & HANDSOME MEN FACE WASH 100G',
    'TRESEMME SHAMPOO HAIR FALL DEFENSE 580ML',

    'GLOW & LOVELY FACIAL MST MLTVIT CRM 100G',
    'CLEAR SHAMPOO COMPLETE ACTVE CARE 90ML',
    'GLOW & LOVELY FACIAL MST MLTVIT CRM 80G',
    'DOVE SHAMPOO NOURISHING OIL 340ML',
    'DOVE SHAMPOO ENVIRONMENTAL DEFENSE 340ML',

    'GLOW & HANDSOME MEN FCE MSTRSR CREAM 50G',
    'DOVE SHAMPOO HEALTHY GROWTH 340ML',
    'PEPSODENT TOOTHPASTE GERMICHECK 20G',
    'GLOW&LOVELY H&B CLR MNGE BODY MILK 100ML',
    'VASELINE LOTION HLTHY WHTE 400ML'
]

# read visits
for file in glob(path + "*.xls"):
    file_name = os.path.basename(file)
    print("Reading data from: " + file_name)
    try: 
        df = pd.read_excel(open(file, "rb"), sheet_name="TradeReturnOutletRevised", header=8, index_col=None)
        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
        row_count = row_count + df.shape[0]
        visit_df = visit_df.append(df)
    except: error_files.append(file_name)
        
# correct columns
cols_to_stay = ["Outlet Code", "Outlet Name", "Channel", "Route", "SSO", "SKUCode", "SKU Description", "Pack Size", "Child Reson Desc", "Qty Ctn", "Qty PC", "Trade Price Per CTN", "Value ", "Secondary Sales", "Percentange Of Damage Against Secondary", "Company Code"]
visit_cols = visit_df.columns.tolist()
for c in cols_to_stay: 
    if c not in visit_cols: 
        visit_df[c] = None
        
# report
print("\nRows to be appended: " + str(row_count))
print("Rows appended: " + str(visit_df.shape[0]))
for e in error_files: print("File couldn't be appended: " + e)
    
# join
print("\nCalculating results...")
qry = '''
select Region, Area, Territory, "Town Name", "Outlet Code", "Outlet Name", Channel, Route, "Route Code", SSO, SKUCode, "SKU Description", "Base Pack", "PS SKU (Y/N)", "NS SKU (Y/N)", "FEFO SKU (Y/N)", "Pack Size", "Child Reson Desc", "Qty Ctn", "Qty PC", "Trade Price Per CTN", "Value ", "Secondary Sales", "Percentange Of Damage Against Secondary", "Company Code"
from 
    visit_df tbl1 
    left join 
    (select "OutletCode" as "Outlet Code", Region, Area, Territory, "TownName" as "Town Name", "RouteCode" as "Route Code"
    from outlet_df
    ) tbl2 using("Outlet Code")
    left join 
    (select Material::string as SKUCode, "Pack size desc." as "Base Pack"
    from ph_df
    ) tbl3 using(SKUCode)
    left join 
    (select sku as "Base Pack", 1 "PS SKU (Y/N)"
    from psns_df 
    where psns_label='PS'
    ) tbl4 using("Base Pack")
    left join 
    (select sku as "Base Pack", 1 "NS SKU (Y/N)"
    from psns_df 
    where psns_label='NS'
    ) tbl5 using("Base Pack")
    left join 
    (select sku as "Base Pack", 1 "FEFO SKU (Y/N)"
    from fefo_df
    ) tbl6 using("Base Pack")
'''
res_df = duckdb.query(qry).df()
file = "C:/Users/Shithi.Maitra/Downloads/Project FEFO/Project FEFO/DAMAGE MIS 22-23/DAMAGE MIS 22-23/Output Files/Output Files/" + "Mother File (Trade Return)_" + path.split("/")[8] + ".xlsx"
res_df.to_excel(file, engine='xlsxwriter', index=False)
file = file.replace("\\", "/")
print("\nFind your results here (" + str(res_df.shape[0]) + " rows): " + file)

# stats
qry = '''
select "PS SKU (Y/N)", "NS SKU (Y/N)", "FEFO SKU (Y/N)", count(*) instances
from res_df 
group by 1, 2, 3
'''
stats_df = duckdb.query(qry).df()
display(stats_df)

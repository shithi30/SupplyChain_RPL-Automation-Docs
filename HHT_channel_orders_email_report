#!/usr/bin/env python
# coding: utf-8

# import
from pathlib import Path
import win32com.client
from win32com.client import Dispatch
import pandas as pd
import duckdb
from pretty_html_table import build_table
import random
from datetime import datetime
import time

# fetch HHT File
def fetch_hht_file(file): 

    # output folder
    output_dir = Path.cwd() / 'HHT Files'
    output_dir.mkdir(parents=True, exist_ok=True)

    # outlook inbox
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    inbox = outlook.Folders.Item(1).Folders['Kader Bhai']
    # inbox = outlook.Folders('shithi.maitra@unilever.com').Folders('Inbox').Folders('Kader Bhai')

    # emails
    messages = inbox.Items
    for message in reversed(messages): 

        # attachments
        attachments = message.Attachments
        for attachment in attachments:
            
            # reports
            filename = attachment.FileName
            if file in filename: 
                print("Found: " + filename)
                attachment.SaveAsFile(output_dir / filename) 
                return

# download
yester_day = duckdb.query('''select strftime(current_date-1, '%d %b %Y') dt''').df()['dt'].tolist()[0]
filename = 'Secondary order_' + yester_day
fetch_hht_file(filename)

# input
ip_df = pd.read_excel(open("C:/Users/Shithi.Maitra/Unilever Codes/Ad Hoc/HHT Files/" + filename + ".xlsx", "rb"), sheet_name="Sheet1", header=0, index_col=None)
ip_df = ip_df[['Classification', 'Basepack', 'Category description', 'Town', 'Company', 'V3 Order HHT']]
ip_df.columns = ['cls', 'bp', 'cat', 'town', 'company', 'hht_order_qty_cs']
ip_df = duckdb.query('''select * from ip_df where hht_order_qty_cs>0''').df()
display(ip_df)

## analysis
# basepack
qry = '''
select 
    bp basepack, 
    sum(hht_order_qty_cs) "hht_order_qty_cs", 
    count(town) "hht_order_town_count", 
    string_agg(town, ', ') "hht_order_towns"
from ip_df 
group by 1 
order by 2 desc
'''
bp_df = duckdb.query(qry).df()
# class
qry = '''select cls "class", sum(hht_order_qty_cs) "hht_order_qty_cs" from ip_df group by 1 order by 2 desc'''
cls_df = duckdb.query(qry).df()
# category
qry = '''select cat category, sum(hht_order_qty_cs) "hht_order_qty_cs" from ip_df group by 1 order by 2 desc'''
cat_df = duckdb.query(qry).df()
# town
qry = '''select town, sum(hht_order_qty_cs) "hht_order_qty_cs" from ip_df group by 1 order by 2 desc'''
town_df = duckdb.query(qry).df()
# company
qry = '''select company, sum(hht_order_qty_cs) "hht_order_qty_cs" from ip_df group by 1 order by 2 desc'''
company_df = duckdb.query(qry).df()

# store
with pd.ExcelWriter("C:/Users/Shithi.Maitra/Downloads/hht_daily_" + yester_day + ".xlsx") as writer:
    ip_df.to_excel(writer, sheet_name="Full", index=False)
    bp_df.to_excel(writer, sheet_name="Basepack", index=False)
    cls_df.to_excel(writer, sheet_name="Class", index=False)
    cat_df.to_excel(writer, sheet_name="Category", index=False)
    company_df.to_excel(writer, sheet_name="Company", index=False)

# email analysis
total_hht_ord = duckdb.query('''select sum("hht_order_qty_cs") hht_order_qty_cs from ip_df''').df()['hht_order_qty_cs'].tolist()[0]
qry = '''
select 
    bp Basepack, 
    sum(hht_order_qty_cs) "HHT Order Qty (CS)", 
    concat(round(sum("hht_order_qty_cs")*100.00/''' + str(total_hht_ord) + ''', 2), '%')  "HHT Order Qty (CS) %",
    count(town) "HHT Order Town Count", 
    case
        when length(string_agg(town, ', '))>40 then concat(left(string_agg(town, ', '), 40), ' ...')
        else string_agg(town, ', ')
    end "HHT Order Towns"
from ip_df 
group by 1 
order by 2 desc
limit 7
'''
email_df = duckdb.query(qry).df()
display(email_df)

# email
ol = win32com.client.Dispatch("outlook.application")
olmailitem = 0x0
newmail = ol.CreateItem(olmailitem)

# subject, recipients
newmail.Subject = 'Daily HHT Distribution'
# newmail.To = 'shithi.maitra@unilever.com'
newmail.To = 'mehedi.asif@unilever.com'
newmail.CC = 'sajeed.jahangir@unilever.com; sanzana.tabassum@unilever.com; anika.hasan@unilever.com; hasib.farabi@unilever.com; asif.rezwan@unilever.com; md.ahsan-habib@unilever.com'

# body
newmail.HTMLbody = f'''
Dear concern,<br><br>
Please find analyses of yesterday's HHT orders (total <b>''' + str(int(total_hht_ord)) + '''</b> cases) attached, in different cuts. Given below is a BP-wise summary (top-<b>07</b>).
''' + build_table(email_df, random.choice(['green_light', 'red_light', 'blue_light', 'grey_light', 'orange_light']), font_size='11px', text_align='left') + '''
More enhancements may be added to the analysis eventually. This is an auto email via <i>win32com</i>.<br><br>
Thanks,<br>
Shithi Maitra<br>
Asst. Manager, Cust. Service Excellence<br>
Unilever BD Ltd.<br>
'''

# attachment(s) 
folder = "C:/Users/Shithi.Maitra/Downloads/"
filename = folder + "hht_daily_" + yester_day + ".xlsx"
newmail.Attachments.Add(filename)

# send
newmail.Send()






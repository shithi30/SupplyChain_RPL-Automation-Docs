"""
Replenishment Report
--------------------
This is a Module with all particulars for generating replenishment for populating SNC.

.. note::
    Rescripting of selective portions are underway.
"""

import pandas as pd
import sys
import win32com.client
import glob as glob
import re
from datetime import date, timedelta
import numpy as np
import time
import shutil
from openpyxl import load_workbook
import xlsxwriter

class Replenishment():
    """
    This class contains all the procedures for generating daily replenishment report.

    .. warning::
        This document just encircles the process of generating the report. Assistive tasks like: downloading, processing of inputs, auto mailing are not within this scope.
    """

    def read_master_inputs(self):
        """
        This function reads the master files (in Excel format) needed to generate replenishment.
        The master files needed are:

        * NMSM, ABC, Winter SKUs' list
        * PDP master
        * UBL, UCL master files
        * PH file
        * Run rate master
        * Cover days master file
        """
        # path
        mpath = r'C:\Users\Shithi.Maitra\Replenishment-report\Master'
        glob.glob(mpath + '\*.xls*')

        # read
        m_nmsm = pd.read_excel(glob.glob(mpath + r'\NMSM*.xlsx')[0],
                               usecols=["Basepack Code", "NMSM TAG", "Customer Code", "Company"])

        m_pdp = pd.read_excel(glob.glob(mpath + r'\PDP*.xlsx')[0],
                              usecols=['Customer Code', 'Region', 'AREA', 'Town', 'Depot Name', 'SAT', 'SUN', 'MON',
                                       'TUE', 'WED', 'THU', 'FRI', 'Company'])

        UBL_m = pd.read_excel(glob.glob(mpath + r'\UBL_Material*.xls*')[0], header=2,
                              usecols=['MaterialCode', 'Price', 'CategoryDescription'])
        UBL_m = UBL_m.loc[~(UBL_m['CategoryDescription'] == 'FUNCTIONAL NUTRITION')]
        UBL_m = UBL_m.rename(columns={"MaterialCode": "Sku code"})
        UBL_m["Company"] = "UBL"

        UCL_m = pd.read_excel(glob.glob(mpath + r'\UCL_Material*.xls*')[0], header=2,
                              usecols=['MaterialCode', 'Price', 'CategoryDescription'])
        UCL_m = UCL_m.loc[(UCL_m['CategoryDescription'] == 'FUNCTIONAL NUTRITION')]
        UCL_m = UCL_m.rename(columns={"MaterialCode": "Sku code"})
        UCL_m["Company"] = "UCL"

        m_full = pd.concat([UBL_m, UCL_m], ignore_index=True)

        Pattern = re.compile("Selling.*")
        xl = pd.ExcelFile(glob.glob(mpath + "\PH*")[0])
        for sheet_name in xl.sheet_names:
            if re.findall(Pattern, sheet_name):
                m_ph = pd.read_excel(glob.glob(mpath + r'\PH*.xlsx')[0], sheet_name=sheet_name,
                                     usecols=['Category description', 'Pack size', 'Pack size desc.', 'Material',
                                              'Company'])

        m_rr = pd.read_excel(glob.glob(mpath + r'\RR*.xlsx')[0])

        m_abc = pd.read_excel(glob.glob(mpath + r'\SKU_Classification*.xlsx')[0],
                              usecols=['Customer Code', "Basepack Code", 'Classification'])

        m_winter = pd.read_excel(glob.glob(mpath + r'\Winter*.xlsx')[0], usecols=["Basepack Code", 'Winter'])

        c_days = pd.read_excel(glob.glob(mpath + r'\rptCoverdaysUpload*.xlsx')[0],
                               usecols=["Location", 'Basepack', 'CoverDays'])
        c_days = c_days.rename(
            columns={"Location": "Customer Code", "Basepack": "Basepack Code", 'CoverDays': "New_CoverDays"})

    def read_daily_inputs():
        """
        The master files are rarely changed. The frequency of change may be once a month.
        Replenishment takes some rolling data as inputs, they are:

        * Yesterday's replenishment report, upon which today's replenishment builds.
        * Inventory status reports, for both UBL and UCL, which must be updated.
        * Daily locks, containing limits on upper/lower boundaries, for generating upper/lower locks.
        * DR file, for some calculated fields.
        * Fund file, to ensure Unilever is allocating only to non-defaulters with above tk.200,000 balance.
        * Proposed order file
        """

        i_path = r'C:\Users\Shithi.Maitra\Replenishment-report\input'
        glob.glob(i_path + '\*.xls*')

        i_rpl = pd.read_excel(glob.glob(i_path + r'\RPL*.xls*')[0])
        i_rpl = i_rpl.drop_duplicates(subset=['Customer code', 'Sku code'], keep='first')
        i_rpl.to_excel(r"C:\Users\Shithi.Maitra\Replenishment-report\Report\i_rpltest_.xlsx", index=False)

        i_invb = pd.read_excel(glob.glob(i_path + r'\UBL_Inventory*.xls*')[0])
        i_invb.drop_duplicates(subset="Sender", keep="first", inplace=True)

        i_invc = pd.read_excel(glob.glob(i_path + r'\UCL_Inventory*.xls*')[0])
        i_invc.drop_duplicates(subset="Sender", keep="first", inplace=True)

        i_inv = pd.concat([i_invb, i_invc], ignore_index=True)
        i_inv = i_inv.rename(
            columns={"IDoc/MsgId": "Company Code", "Sender": "Customer Code", "Status": "Customer Inventory Status"})
        i_inv["Company"] = np.where(i_inv["Company Code"] == 1532, "UBL", "UCL")

        i_lock = pd.read_excel(glob.glob(i_path + r'\Lock*.xls*')[0],
                               usecols=['Ship to', 'Product', 'Tolerance Lower band', 'Tolerance Upper band'])

        i_dr = pd.read_excel(glob.glob(i_path + r'\DR*.xls*')[0])  # edit
        i_dr.drop(i_dr.index[:2], inplace=True)

        i_fundb = pd.read_excel(glob.glob(i_path + r'\UBL-Fund*.xls*')[0], header=3, sheet_name="Fund",
                                usecols=["Distributor", "Comment"])
        i_fundb["Company"] = "UBL"
        i_fundc = pd.read_excel(glob.glob(i_path + r'\UCL-Fund*.xls*')[0], header=3, usecols=["Distributor", "Comment"])
        i_fundc["Company"] = "UCL"
        i_fund = pd.concat([i_fundb, i_fundc], ignore_index=True)

        p_order = pd.read_excel(glob.glob(i_path + r'\Proposed_order*.xls*')[0],
                                usecols=['Destination Location', 'Product Number', 'Proposed Quantity'])
        p_order = p_order.rename(columns={"Destination Location": "Customer Code", "Product Number": "Sku code",
                                          'Proposed Quantity': 'New_Proposed Quantity'})
    def clean_daily_inputs(self):
        """
        The daily inputs files' data must be morphed into a form consistent with the replenishment ethos.
        Such measures include:

        * Renaming of columns for consistent keys for joining/merging.
        * Imputation of null values by 0s, to avoid errors/complexities.
        * Removal/aggregation of duplicate entries.
        * Transformation of values for the sake of calculation: changing of data types, rearrangement of columns.
        """

        i_rpl = i_rpl.rename(columns={"Customer code": "Customer Code"})
        m_ph = m_ph.rename(
            columns={"Material": "Sku code", "Pack size desc.": "Basepack", "Pack size": "Basepack Code"})
        i_lock = i_lock.rename(columns={"Ship to": "Customer Code", "Product": "Sku code"})
        i_fund = i_fund.rename(columns={"Distributor": "Customer Code", "Comment": "Defaulter list"})

        i_dr.fillna(0, inplace=True)
        i_dr = i_dr.rename(
            columns={"Unnamed: 2": "Customer Code", "Unnamed: 3": "Basepack", "Unnamed: 4": "Basepack Code"})
        i_dr["Total order"] = i_dr.filter(like="Final Customer").sum(axis=1)
        i_dr["Total Dispatched"] = i_dr.filter(like="Dispatched Qty").sum(axis=1)
        i_dr["DR%"] = (i_dr["Total Dispatched"] / i_dr["Total order"]) * 100
        i_dr = i_dr.rename(columns={"DR%": "Past week DR%"})
        i_dr = i_dr[['Customer Code', 'Basepack Code', 'Past week DR%']]
        i_dr = i_dr[['Customer Code', 'Basepack Code', 'Past week DR%']].groupby(
            by=['Customer Code', 'Basepack Code']).sum()
        i_dr['Customer Code'] = i_dr['Customer Code'].astype("int64")
        i_dr.reset_index(inplace=True)


    def prepare_rpl_report(self):
        """
        This function is responsible for generating the full RPL report.
        The report consists of 56 fields that can be categorized as follows:

        * **Day:** 'Date', 'OPD'
        
        * **Customer:** 'Customer Code', 'Customer name', 'Customer Inventory Status', 'Defaulter list', 'Town', 'Depot Name', 'Region', 'AREA'
        
        * **Product:** 'Company', 'Category description', 'Winter', or, 'Classification', 'NMSM TAG', 'Sku code', 'Sku description', 'Basepack', 'Basepack Code', 'Price / UOM'
        
        * **Stock, Norm, Proposed, Cover Days:** 'Norm qty', 'Norm_Value', 'Cover days', 'Stock on hand', 'Stock on hand_Value', 'In transit', 'In transit_Value', 'Open order', 'Open order_Value', 'Stock cover in days', 'Proposed qty', 'PPO value',

        * **Tolerance, Minmax, Capping, MSTN:** 'Tolerance Upper band','Tolerance Lower band', 'Max saved value','Min saved value', 'Min of Phy stock capping norm', 'Min of Total stock capping norm', 'Phy MSTN','Total MSTN'

        * **Historic Sales:** 'Past week DR%', 'Perday sales qty', 'Perday sales Value'

        * **OOS:** 'OOS qty','OOS Value', 'PSL', 'PSL-Considering Coverdays', 'ZERO', '<1 day', '<2 day', '<3 day', 'SC-RR', 'ZERO Stock', 'Stock<1 day','Stock<2 day', 'Stock<3 day', 'Loss reason'
        """


        # # Prepare Replenishment_report

        # In[61]:

        i_rpl = pd.merge(i_rpl, m_ph, on=["Sku code"], how="left")

        # In[62]:

        i_rpl = pd.merge(i_rpl, p_order, on=["Customer Code", "Sku code"], how="left")

        # In[63]:

        i_rpl['Proposed qty'] = np.where(i_rpl['New_Proposed Quantity'].isna(), i_rpl['Proposed qty'],
                                         i_rpl['New_Proposed Quantity'])

        # In[64]:

        i_rpl = pd.merge(i_rpl, c_days, on=["Customer Code", "Basepack Code"], how="left")

        # In[65]:

        i_rpl['Cover days'] = np.where(i_rpl['New_CoverDays'].isna(), i_rpl['Cover days'], i_rpl['New_CoverDays'])

        # In[66]:

        i_rpl_new_cal = i_rpl[
            ['Customer Code', 'Customer name', 'Sku code', 'Sku description', 'Proposed qty', 'Price / UOM', 'Norm qty',
             'Cover days',
             'Stock on hand', 'In transit', 'Open order', 'PPO value', 'Category description', 'Basepack Code',
             'Basepack', 'Company']]

        # In[67]:

        i_rpl_new_cal.columns

        # In[68]:

        i_rpl_new_cal['Correct_Proposed qty'] = i_rpl_new_cal['Norm qty'] - i_rpl_new_cal['Stock on hand'] - \
                                                i_rpl_new_cal['In transit'] - i_rpl_new_cal['Open order']

        # In[69]:

        i_rpl_new_cal['Correct_Proposed qty'] = np.where(i_rpl_new_cal['Correct_Proposed qty'] <= 0, 0,
                                                         i_rpl_new_cal['Correct_Proposed qty'])

        # In[70]:

        i_rpl_new_cal['Gap'] = i_rpl_new_cal['Correct_Proposed qty'] - i_rpl_new_cal['Proposed qty']

        # In[71]:

        Gap_Proposed_qty = i_rpl_new_cal[
            ['Customer Code', 'Sku code', 'Sku description', 'Proposed qty', 'Price / UOM', 'Norm qty', 'Cover days',
             'Stock on hand', 'In transit', 'Open order', 'PPO value', 'Category description', 'Basepack Code',
             'Basepack',
             'Company', 'Correct_Proposed qty', 'Gap']]

        # In[72]:

        Gap_Proposed_qty = Gap_Proposed_qty[~(Gap_Proposed_qty['Gap'] == 0)]

        # In[73]:

        Stock_in_hand = i_rpl_new_cal[
            ['Customer Code', 'Sku code', 'Sku description', 'Proposed qty', 'Price / UOM', 'Norm qty', 'Cover days',
             'Stock on hand', 'In transit', 'Open order', 'PPO value', 'Category description', 'Basepack Code',
             'Basepack',
             'Company', 'Correct_Proposed qty', 'Gap']]

        # In[74]:

        Stock_in_hand = Stock_in_hand[Stock_in_hand['Stock on hand'] == 0]

        # In[75]:

        i_rpl = i_rpl[
            ['Customer Code', 'Customer name', 'Sku code', 'Sku description', 'Proposed qty', 'Price / UOM', 'Norm qty',
             'Cover days',
             'Stock on hand', 'In transit', 'Open order', 'PPO value', 'Category description', 'Basepack Code',
             'Basepack', 'Company']]

        # In[76]:

        price_check = i_rpl.loc[(i_rpl['Price / UOM'] == 0)]

        # In[77]:

        t = pd.to_datetime('today').strftime('%d %b %Y')

        # In[78]:

        price_check.to_excel(r"C:\Users\Abdul.Kader\Replenishment-report\Report\price_check_" + t + ".xlsx",
                             index=False)

        # In[79]:

        i_rpl = pd.merge(i_rpl, m_full[["Sku code", "Price"]], on=["Sku code"], how="left")

        # In[80]:

        i_rpl['Price / UOM'] = np.where(i_rpl['Price / UOM'] == 0, i_rpl['Price'], i_rpl['Price / UOM'])

        # In[81]:

        i_rpl['PPO value'] = np.where(i_rpl['PPO value'] == 0, i_rpl['Price'] * i_rpl['Proposed qty'],
                                      i_rpl['PPO value'])

        # In[82]:

        i_rpl['PPO value'] = np.where(i_rpl['PPO value'].isna(), 0, i_rpl['PPO value'])

        # In[83]:

        # i_rpl.to_excel(r"C:\Users\Abdul.Kader\Replenishment-report\Report\i_rpl_"+t+".xlsx", index = False)

        # In[84]:

        i_rpl = pd.merge(i_rpl, i_inv[["Customer Code", "Company", "Customer Inventory Status"]],
                         on=["Customer Code", "Company"], how="left")

        # In[85]:

        # i_rpl.fillna({'Dispatched Qty.':0,'Dispatched Rate %  (Tactical Measure)': 0}, inplace=True)

        # In[86]:

        i_rpl["Customer Inventory Status"].fillna("Not Update", inplace=True)

        # In[87]:

        i_rpl["Customer Inventory Status"] = np.where(i_rpl["Customer Inventory Status"] == "Not Update", "Not Update",
                                                      "Update")

        # In[88]:

        # i_rpl["Customer Inventory Status"] = i_rpl["Company"].apply(lambda x: "Updated" if x in set(i_inv["Company"].unique()) else "Not updated")

        # In[89]:

        i_rpl = pd.merge(i_rpl, i_lock, on=["Customer Code", "Sku code"], how="left")

        # In[90]:

        i_rpl = pd.merge(i_rpl, i_dr, on=["Customer Code", "Basepack Code"], how="left")

        # In[91]:

        i_rpl["Past week DR%"] = i_rpl["Past week DR%"].fillna("No-Order")

        # In[92]:

        i_rpl = pd.merge(i_rpl, i_fund[["Customer Code", "Defaulter list", "Company"]], on=["Customer Code", "Company"],
                         how="left")

        # In[93]:

        i_rpl["Defaulter list"].fillna("Active", inplace=True)

        # In[94]:

        i_rpl.shape

        # In[95]:

        i_rpl["Defaulter list"] = np.where(i_rpl["Defaulter list"] == "Active", "Active", "Defaulter")

        # In[96]:

        # i_rpl["Defaulter list"] = i_rpl["Defaulter list"].apply(lambda x:"Active" if x ==0 else "Defaulter")

        # In[97]:

        rdate2 = input('Insert PDP Date: ')
        rdate = rdate2

        # In[98]:

        rdate = pd.to_datetime(rdate).strftime("%a").upper()

        # In[99]:

        m_pdp.columns

        # In[100]:

        i_rpl = pd.merge(i_rpl, m_pdp[
            ["Customer Code", "Region", "AREA", "Town", "Depot Name", "Company", m_pdp.filter(like=rdate).columns[0]]],
                         on=["Customer Code", "Company"], how="left")

        # In[101]:

        # i_rpl = pd.merge (i_rpl, m_pdp[["Customer Code","Region","Town","Depot Name","Company",m_pdp.filter(like = rdate).columns[0]]], on= ["Customer Code","Company"], how= "left" )

        # In[102]:

        i_rpl = i_rpl.rename(columns={rdate: "OPD"})

        # In[103]:

        i_rpl = pd.merge(i_rpl, m_rr[["Customer Code", "Basepack Code", "L6 daily RR"]],
                         on=["Customer Code", "Basepack Code"], how="left")

        # In[104]:

        lock_check = i_rpl[
            ['Customer Code', 'Sku code', 'Sku description', 'Category description', 'Basepack Code', 'Basepack',
             "Town", 'Company', 'Tolerance Lower band', 'Tolerance Upper band']]

        # In[105]:

        lock_check = lock_check.loc[~(lock_check['Tolerance Lower band'] >= 0)]

        # In[106]:

        lock_check.to_excel(r"C:\Users\Abdul.Kader\Replenishment-report\Report\lock_check_" + rdate2 + ".xlsx",
                            index=False)

        # In[107]:

        i_rpl.columns

        # In[108]:

        i_rpl["Tolerance Lower band"] = np.where(((i_rpl["Company"] == "UCL") & (i_rpl["Tolerance Lower band"].isna())),
                                                 50, i_rpl["Tolerance Lower band"])

        # In[109]:

        i_rpl["Tolerance Upper band"] = np.where(((i_rpl["Company"] == "UCL") & (i_rpl["Tolerance Upper band"].isna())),
                                                 100, i_rpl["Tolerance Upper band"])

        # In[110]:

        i_rpl["Tolerance Lower band"] = np.where(((i_rpl["Company"] == "UBL") & (i_rpl["Tolerance Lower band"].isna())),
                                                 50, i_rpl["Tolerance Lower band"])

        # In[111]:

        i_rpl["Tolerance Upper band"] = np.where(((i_rpl["Company"] == "UBL") & (i_rpl["Tolerance Upper band"].isna())),
                                                 100, i_rpl["Tolerance Upper band"])

        # In[112]:

        # i_rpl.to_excel(r"C:\Users\Abdul.Kader\Replenishment-report\Report\i_rpl_"+rdate2+".xlsx", index = False)

        # In[113]:

        # i_rpl.to_excel(r"C:\Users\Abdul.Kader\Replenishment-report\Report\i_rpltest_"+rdate2+".xlsx", index = False)

        # In[114]:

        i_rpl["L6 daily RR"].fillna(0, inplace=True)

        # In[115]:

        i_rpl = pd.merge(i_rpl, m_winter, on="Basepack Code", how="left")

        # In[116]:

        i_rpl = pd.merge(i_rpl, m_abc, on=["Customer Code", "Basepack Code"], how="left")

        # In[117]:

        i_rpl["Winter"].fillna(0, inplace=True)

        # In[118]:

        # i_rpl.assign(Class=lambda x: ("Winter" if x["Winter"] !=0 else x["Classification"]))

        # In[119]:

        i_rpl["Classification"] = i_rpl["Winter"].where(i_rpl["Winter"] == "Winter", i_rpl["Classification"])

        # In[120]:

        i_rpl["Classification"].fillna("D", inplace=True)

        # In[121]:

        i_rpl = pd.merge(i_rpl, m_nmsm, on=["Customer Code", "Basepack Code", "Company"], how="left")

        # In[122]:

        i_rpl['NMSM TAG'] = np.where(i_rpl['NMSM TAG'].isna(), "No NMSM TAG", i_rpl['NMSM TAG'])

        # In[123]:

        i_rpl["Stock on hand_Value"] = i_rpl["Price / UOM"] * i_rpl["Stock on hand"]

        # In[124]:

        i_rpl["In transit_Value"] = i_rpl["Price / UOM"] * i_rpl["In transit"]

        # In[125]:

        i_rpl["Open order_Value"] = i_rpl["Price / UOM"] * i_rpl["Open order"]

        # In[126]:

        i_rpl["Max saved value"] = (1 + (i_rpl["Tolerance Upper band"] / 100)) * i_rpl["Price / UOM"] * i_rpl[
            "Proposed qty"]

        # In[127]:

        i_rpl["Min saved value"] = np.ceil((1 - (i_rpl["Tolerance Lower band"] / 100)) * i_rpl["Proposed qty"]) * i_rpl[
            "Price / UOM"]

        # In[128]:

        i_rpl['Min of Phy stock capping norm'] = i_rpl[['Norm qty', 'Stock on hand']].min(axis=1)

        # In[129]:

        i_rpl["Phy MSTN"] = (i_rpl["Min of Phy stock capping norm"] / i_rpl["Norm qty"]) * 100

        # In[130]:

        i_rpl["Min of Total stock capping norm"] = round(
            np.minimum((i_rpl["Stock on hand"] + i_rpl["In transit"]), i_rpl['Norm qty']))

        # In[131]:

        # i_rpl["Soh IT"] = i_rpl["Stock on hand"] + i_rpl["In transit"]

        # In[132]:

        # i_rpl["Min of Total stock capping norm"]=i_rpl[['Norm qty','Soh IT']].min(axis=1)

        # In[133]:

        i_rpl["Total MSTN"] = (i_rpl["Min of Total stock capping norm"] / i_rpl["Norm qty"]) * 100

        # In[134]:

        i_rpl = i_rpl.rename(columns={"L6 daily RR": "Perday sales qty"})

        # In[135]:

        i_rpl["Perday sales Value"] = i_rpl["Perday sales qty"] * i_rpl["Price / UOM"]

        # In[136]:

        i_rpl["OOS qty"] = np.where(i_rpl["Stock on hand"] <= 0, i_rpl["Norm qty"], 0)

        # In[137]:

        i_rpl["OOS Value"] = i_rpl["OOS qty"] * i_rpl["Price / UOM"]

        # In[138]:

        i_rpl["PSL"] = np.where(i_rpl["Stock on hand"] == 0, i_rpl["Perday sales Value"], 0)

        # In[139]:

        i_rpl["Norm_Value"] = i_rpl["Norm qty"] * i_rpl["Price / UOM"]

        # In[140]:

        i_rpl["Past week DR% Help"] = np.where(i_rpl["Past week DR%"] == "No-Order", 9999, i_rpl["Past week DR%"])

        # In[141]:

        i_rpl["Loss reason"] = np.where(i_rpl["Stock on hand"] < i_rpl["Norm qty"],
                                        np.where(i_rpl["Defaulter list"] == "Defaulter", "Defaulter",
                                                 np.where(((i_rpl["Past week DR% Help"] == 9999) & (
                                                             i_rpl["In transit"] <= 0)), "No-Order from Customer",
                                                          np.where((i_rpl["Past week DR% Help"] < 85), "Supply Issue",
                                                                   np.where((i_rpl["In transit"] > 0),
                                                                            "Stock in transit",
                                                                            "No-Order from Customer")
                                                                   )
                                                          )
                                                 ), "Stock available"
                                        )

        # In[142]:

        i_rpl["Stock cover in days"] = i_rpl["Stock on hand"] / (i_rpl["Norm qty"] / i_rpl["Cover days"])

        # In[143]:

        i_rpl["Stock cover in days"].fillna(0, inplace=True)

        # In[144]:

        i_rpl["PSL-Considering Coverdays"] = np.where(i_rpl["Stock on hand"] > i_rpl["Perday sales qty"], 0,
                                                      i_rpl["Perday sales Value"] - i_rpl["Stock on hand_Value"])

        # In[145]:

        i_rpl["ZERO"] = np.where(i_rpl["Stock cover in days"] == 0, i_rpl["PSL-Considering Coverdays"], 0)

        # In[146]:

        i_rpl["<1 day"] = np.where((i_rpl["Stock cover in days"] <= 1) & (i_rpl["Stock cover in days"] > 0),
                                   i_rpl["PSL-Considering Coverdays"], 0)

        # In[147]:

        i_rpl["<2 day"] = np.where((i_rpl["Stock cover in days"] > 1) & (i_rpl["Stock cover in days"] <= 2),
                                   i_rpl["PSL-Considering Coverdays"], 0)

        # In[148]:

        i_rpl["<3 day"] = np.where((i_rpl["Stock cover in days"] > 2) & (i_rpl["Stock cover in days"] <= 3),
                                   i_rpl["PSL-Considering Coverdays"], 0)

        # In[149]:

        i_rpl["SC-RR"] = i_rpl["Stock on hand_Value"] / i_rpl["Perday sales Value"]

        # In[150]:

        i_rpl["SC-RR"].fillna(0, inplace=True)

        # In[151]:

        i_rpl["ZERO Stock"] = np.where(i_rpl["SC-RR"] == 0, i_rpl["Perday sales Value"], 0)

        # In[152]:

        i_rpl["Stock Day1"] = np.where((i_rpl["SC-RR"] <= 1) & (i_rpl["SC-RR"] > 0),
                                       i_rpl["Perday sales Value"] - i_rpl["Stock on hand_Value"], 0)

        # In[153]:

        i_rpl["Stock<1 day"] = np.where(i_rpl["SC-RR"] == 0, (i_rpl["Perday sales Value"]), i_rpl["Stock Day1"])

        # In[154]:

        conditions = [
            (i_rpl["SC-RR"] == 0),
            (i_rpl["Stock<1 day"] > 0),
            ((i_rpl["SC-RR"] > 1) & (i_rpl["SC-RR"] <= 2)),

        ]

        values = [
            (2 * i_rpl["Perday sales Value"]),
            (i_rpl["Stock<1 day"] + i_rpl["Perday sales Value"]),
            ((i_rpl["Perday sales Value"] * 2) - i_rpl["Stock on hand_Value"])

        ]

        i_rpl["Stock<2 day"] = np.select(conditions, values)

        # In[155]:

        conditions = [
            (i_rpl["SC-RR"] == 0),
            (i_rpl["Stock<2 day"] > 0),
            ((i_rpl["SC-RR"] > 2) & (i_rpl["SC-RR"] <= 3)),

        ]

        values = [
            (3 * i_rpl["Perday sales Value"]),
            (i_rpl["Stock<2 day"] + i_rpl["Perday sales Value"]),
            ((i_rpl["Perday sales Value"] * 3) - i_rpl["Stock on hand_Value"])

        ]

        i_rpl["Stock<3 day"] = np.select(conditions, values)

        # In[156]:

        # i_rpl_col = pd.DataFrame(i_rpl.columns,columns = ["Columns Name"])

        # In[157]:

        # i_rpl_col["Serial"] = list(i_rpl_col.index)

        # In[158]:

        # i_rpl_Serial = pd.read_excel(r"C:\Users\Abdul.Kader\Replenishment-report\i_rpl_col.xlsx")

        # In[159]:

        # i_rpl = i_rpl[list(i_rpl_Serial["Columns Name"])]

        # In[160]:

        i_rpl['Date'] = rdate2

        # In[161]:

        i_rpl = i_rpl[
            ['Date', 'OPD', 'Customer Inventory Status', 'Classification', 'Customer Code', 'Customer name', 'Sku code',
             'Sku description',
             'Basepack', 'Basepack Code', 'Category description', 'Proposed qty', 'Price / UOM', 'Norm qty',
             'Cover days', 'Stock on hand', 'In transit',
             'Open order', 'PPO value', 'Town', 'Depot Name', 'Region', 'AREA', 'Stock on hand_Value',
             'In transit_Value', 'Open order_Value',
             'Tolerance Upper band', 'Tolerance Lower band', 'Max saved value', 'Min saved value',
             'Min of Phy stock capping norm',
             'Min of Total stock capping norm', 'Phy MSTN', 'Total MSTN', 'Perday sales qty', 'Perday sales Value',
             'OOS qty', 'OOS Value',
             'PSL', 'Defaulter list', 'Past week DR%', 'NMSM TAG', 'Company', 'Norm_Value', 'Loss reason',
             'Stock cover in days',
             'PSL-Considering Coverdays', 'ZERO', '<1 day', '<2 day', '<3 day', 'SC-RR', 'ZERO Stock', 'Stock<1 day',
             'Stock<2 day', 'Stock<3 day']]

        # In[162]:

        i_rpl[i_rpl.select_dtypes(include=["float64"]).columns] = i_rpl.select_dtypes(include=["float64"]).round(2)

    def pivot(self):
        """
        This function is responsible for generating all pivots that go in the email body and other sheets.
        """

        # In[163]:

        classification_Summary = (pd.pivot_table(i_rpl, index=['Classification'],
                                                 values=['PPO value', 'Stock on hand_Value', 'In transit_Value',
                                                         'Open order_Value', 'Norm_Value', 'Max saved value',
                                                         'Min saved value'], aggfunc='sum', margins=True,
                                                 margins_name='Total') / 10000000).round(2)

        # In[164]:

        nmsm_tag_Summary = (pd.pivot_table(i_rpl, index=['NMSM TAG'],
                                           values=['PPO value', 'Stock on hand_Value', 'In transit_Value',
                                                   'Open order_Value', 'Norm_Value', 'Max saved value',
                                                   'Min saved value'], aggfunc='sum', margins=True,
                                           margins_name='Total') / 10000000).round(2)

        # In[165]:

        classification_nstm_Summary = (
                    pd.pivot_table(i_rpl, index=['Classification'], values=['PSL', 'Phy MSTN', 'Total MSTN'],
                                   aggfunc='sum', margins=True, margins_name='Total') / 10000000).round(2)

        # In[166]:

        region_nstm_Summary = (
                    pd.pivot_table(i_rpl, index=['Region'], values=['PSL', 'Phy MSTN', 'Total MSTN'], aggfunc='sum',
                                   margins=True, margins_name='Total') / 10000000).round(2)

        # In[167]:

        loss_reason_nstm_Summary = (
                    pd.pivot_table(i_rpl, index=['Loss reason'], values=['PSL', 'Phy MSTN', 'Total MSTN'],
                                   aggfunc='sum', margins=True, margins_name='Total') / 10000000).round(2)

        # In[168]:

        region_loss_reason_nstm_Summary = (
                    pd.pivot_table(i_rpl, index=['Region'], columns=['Loss reason'], values=['PSL'], aggfunc='sum',
                                   margins=True, margins_name='Total') / 10000000).round(2)

        # In[169]:

        classification_day_Summary = (
                    pd.pivot_table(i_rpl, index=['Classification'], values=['ZERO', '<1 day', '<2 day', '<3 day'],
                                   aggfunc='sum', margins=True, margins_name='Total') / 10000000).round(2)

        # In[170]:

        loss_reason_day_Summary = (
                    pd.pivot_table(i_rpl, index=['Loss reason'], values=['ZERO', '<1 day', '<2 day', '<3 day'],
                                   aggfunc='sum', margins=True, margins_name='Total') / 10000000).round(2)

        # In[171]:

        classification_stock_Summary = (pd.pivot_table(i_rpl, index=['Classification'],
                                                       values=['ZERO Stock', 'Stock<1 day', 'Stock<2 day',
                                                               'Stock<3 day'], aggfunc='sum', margins=True,
                                                       margins_name='Total') / 10000000).round(2)

        # In[172]:

        loss_reason_stock_Summary = (pd.pivot_table(i_rpl, index=['Loss reason'],
                                                    values=['ZERO Stock', 'Stock<1 day', 'Stock<2 day', 'Stock<3 day'],
                                                    aggfunc='sum', margins=True,
                                                    margins_name='Total') / 10000000).round(2)

        # In[173]:

        classification_stock_loss_reason_Summary = (pd.pivot_table(i_rpl, index=['Classification', 'Loss reason'],
                                                                   values=['ZERO Stock', 'Stock<1 day', 'Stock<2 day',
                                                                           'Stock<3 day'], aggfunc='sum', margins=True,
                                                                   margins_name='Total') / 10000000).round(2)

        # In[174]:

        region_stock_classification_Summary = (pd.pivot_table(i_rpl, index=['Region', 'Classification'],
                                                              values=['ZERO Stock', 'Stock<1 day', 'Stock<2 day',
                                                                      'Stock<3 day'], aggfunc='sum', margins=True,
                                                              margins_name='Total') / 10000000).round(2)

        # In[175]:

        # book = load_workbook(r"C:\\Users\Abdul.Kader\\Replenishment-report\\Report\\Replenishment Repot_"+rdate2+".xlsx")
        writer = pd.ExcelWriter(
            r"C:\Users\Abdul.Kader\Replenishment-report\Report\Replenishment Repot_" + rdate2 + ".xlsx",
            engine='openpyxl')

        # In[176]:

        classification_Summary.to_excel(writer, "Full summary", startcol=0, startrow=0)
        nmsm_tag_Summary.to_excel(writer, "Full summary", startcol=0, startrow=8)
        classification_nstm_Summary.to_excel(writer, "Full summary", startcol=0, startrow=16)
        # classification_nstm_Summary.to_excel(writer, "Full summary", startcol=0,startrow= 23)
        region_nstm_Summary.to_excel(writer, "Full summary", startcol=0, startrow=32)
        loss_reason_nstm_Summary.to_excel(writer, "Full summary", startcol=0, startrow=49)
        region_loss_reason_nstm_Summary.to_excel(writer, "Full summary", startcol=0, startrow=60)
        classification_day_Summary.to_excel(writer, "Full summary", startcol=0, startrow=79)
        loss_reason_day_Summary.to_excel(writer, "Full summary", startcol=0, startrow=87)
        classification_stock_Summary.to_excel(writer, "Full summary", startcol=0, startrow=96)
        loss_reason_stock_Summary.to_excel(writer, "Full summary", startcol=0, startrow=104)
        classification_stock_loss_reason_Summary.to_excel(writer, "Full summary", startcol=0, startrow=114)
        region_stock_classification_Summary.to_excel(writer, "Full summary", startcol=0, startrow=136)
        i_rpl.to_excel(writer, "Replenishment UBL_UCL", startcol=0, startrow=0, index=False)
        # writer.save()

        writer.close()

        # In[177]:

        inventory = i_rpl[
            ['Date', 'Region', 'AREA', 'OPD', 'Customer Inventory Status', 'Customer Code', 'Town', 'Company']]

        inventory_UBL = inventory[inventory['Company'] == 'UBL']

        inventory_UBL = inventory_UBL.drop_duplicates()

        inventory_UCL = inventory[inventory['Company'] == 'UCL']

        inventory_UCL = inventory_UCL.drop_duplicates()

        writer = pd.ExcelWriter(
            r"C:\Users\Abdul.Kader\Replenishment-report\Report\Inventory_Update_List_UBL_&_UCL_" + rdate2 + ".xlsx",
            engine='openpyxl')

        inventory_UBL.to_excel(writer, "UBL_Inventory", startcol=0, startrow=0, index=False)

        inventory_UCL.to_excel(writer, "UCL_Inventory", startcol=0, startrow=0, index=False)
        # writer.save()

        writer.close()

        # In[178]:

        Gap_Proposed_qty.to_excel(r"C:\Users\Abdul.Kader\Replenishment-report\Report\Gap_Proposed_qty_" + t + ".xlsx",
                                  index=False)

        Stock_in_hand.to_excel(r"C:\Users\Abdul.Kader\Replenishment-report\Report\Stock_in_hand_" + t + ".xlsx",
                               index=False)

        # # UBL Area

        # In[179]:

        ubl = i_rpl[
            ['Date', 'OPD', 'Customer Inventory Status', 'Classification', 'Customer Code', 'Customer name', 'Sku code',
             'Sku description',
             'Basepack', 'Category description', 'Proposed qty', 'Price / UOM', 'Norm qty', 'Cover days',
             'Stock on hand', 'In transit',
             'Open order', 'PPO value', 'Town', 'Depot Name', 'Region', 'AREA', 'Stock on hand_Value',
             'In transit_Value', 'Open order_Value',
             'Tolerance Upper band', 'Tolerance Lower band', 'Max saved value', 'Min saved value',
             'Min of Phy stock capping norm',
             'Min of Total stock capping norm', 'Phy MSTN', 'Total MSTN', 'Perday sales qty', 'Perday sales Value',
             'OOS qty', 'OOS Value',
             'PSL', 'Defaulter list', 'Past week DR%', 'NMSM TAG', 'Company', 'Norm_Value', 'Loss reason',
             'Stock cover in days',
             'PSL-Considering Coverdays', 'ZERO', '<1 day', '<2 day', '<3 day', 'SC-RR', 'ZERO Stock', 'Stock<1 day',
             'Stock<2 day', 'Stock<3 day']]

        # In[180]:

        ubl = ubl[ubl['Company'] == "UBL"]

        # In[181]:

        classification_Summary = (pd.pivot_table(ubl, index=['Classification'],
                                                 values=['PPO value', 'Stock on hand_Value', 'In transit_Value',
                                                         'Open order_Value', 'Norm_Value', 'Max saved value',
                                                         'Min saved value'], aggfunc='sum', margins=True,
                                                 margins_name='Total') / 10000000).round(2)

        nmsm_tag_Summary = (pd.pivot_table(ubl, index=['NMSM TAG'],
                                           values=['PPO value', 'Stock on hand_Value', 'In transit_Value',
                                                   'Open order_Value', 'Norm_Value', 'Max saved value',
                                                   'Min saved value'], aggfunc='sum', margins=True,
                                           margins_name='Total') / 10000000).round(2)

        classification_nstm_Summary = (
                    pd.pivot_table(ubl, index=['Classification'], values=['PSL', 'Phy MSTN', 'Total MSTN'],
                                   aggfunc='sum', margins=True, margins_name='Total') / 10000000).round(2)

        region_nstm_Summary = (
                    pd.pivot_table(ubl, index=['Region'], values=['PSL', 'Phy MSTN', 'Total MSTN'], aggfunc='sum',
                                   margins=True, margins_name='Total') / 10000000).round(2)

        loss_reason_nstm_Summary = (
                    pd.pivot_table(ubl, index=['Loss reason'], values=['PSL', 'Phy MSTN', 'Total MSTN'], aggfunc='sum',
                                   margins=True, margins_name='Total') / 10000000).round(2)

        region_loss_reason_nstm_Summary = (
                    pd.pivot_table(ubl, index=['Region'], columns=['Loss reason'], values=['PSL'], aggfunc='sum',
                                   margins=True, margins_name='Total') / 10000000).round(2)

        classification_day_Summary = (
                    pd.pivot_table(ubl, index=['Classification'], values=['ZERO', '<1 day', '<2 day', '<3 day'],
                                   aggfunc='sum', margins=True, margins_name='Total') / 10000000).round(2)

        loss_reason_day_Summary = (
                    pd.pivot_table(ubl, index=['Loss reason'], values=['ZERO', '<1 day', '<2 day', '<3 day'],
                                   aggfunc='sum', margins=True, margins_name='Total') / 10000000).round(2)

        classification_stock_Summary = (pd.pivot_table(ubl, index=['Classification'],
                                                       values=['ZERO Stock', 'Stock<1 day', 'Stock<2 day',
                                                               'Stock<3 day'], aggfunc='sum', margins=True,
                                                       margins_name='Total') / 10000000).round(2)

        loss_reason_stock_Summary = (pd.pivot_table(ubl, index=['Loss reason'],
                                                    values=['ZERO Stock', 'Stock<1 day', 'Stock<2 day', 'Stock<3 day'],
                                                    aggfunc='sum', margins=True,
                                                    margins_name='Total') / 10000000).round(2)

        classification_stock_loss_reason_Summary = (pd.pivot_table(ubl, index=['Classification', 'Loss reason'],
                                                                   values=['ZERO Stock', 'Stock<1 day', 'Stock<2 day',
                                                                           'Stock<3 day'], aggfunc='sum', margins=True,
                                                                   margins_name='Total') / 10000000).round(2)

        region_stock_classification_Summary = (pd.pivot_table(ubl, index=['Region', 'Classification'],
                                                              values=['ZERO Stock', 'Stock<1 day', 'Stock<2 day',
                                                                      'Stock<3 day'], aggfunc='sum', margins=True,
                                                              margins_name='Total') / 10000000).round(2)


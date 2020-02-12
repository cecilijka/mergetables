# pip install openpyxl
# pip install pandas
# pip install xlrd
# pip install Jinja2
import pandas as pd
import numpy as np
import os

# _______________________________ Uber Eats this week
from sqlalchemy import null

uber = pd.read_csv ('files/UberWeek.csv', usecols=["store_name", "Orders", "Sales", "Unfulfilleds"],
                    encoding="utf-8")
# Read tables with Restaurant number and other info
BKNU = pd.read_excel ("Maps\BK#U.xlsx", usecols=["store_name", "BK#"])
# Merge two tables
merge_uber = pd.merge (uber, BKNU)
merge_uber = merge_uber.astype ({"Sales": int})

# save flirterad file
merge_uber.to_csv (r'C:files/UberWeek!.csv')
clean_uber = pd.read_csv ("files/UberWeek!.csv", usecols=["BK#", "Orders", "Sales", "Unfulfilleds"])
clean_uber = clean_uber.rename (
    columns={'Orders': 'Amount of orders (U)', 'Sales': 'Sales (U)', 'Unfulfilleds': 'Denied orders (U)'})
# save flirterad file
clean_uber.to_csv (r'C:files/UberWeek!.csv')

newtable = pd.read_excel ('files/NEWTABLE.xlsx')

pd.set_option ('display.max_columns', None)

mergeU = pd.merge (clean_uber, newtable, on='BK#', how='left')

# save flirterad file
mergeU.to_excel (r'C:files/NEWTABLE!.xlsx')

# _______________________________ Uber Eats last week
uber_last_week = pd.read_csv ('files/UberLastWeek.csv', usecols=["store_name", "Orders", "Sales", "Unfulfilleds"],
                              encoding="utf-8")
# Read tables with Restaurant number and other info
BKNU_last = pd.read_excel ("maps\BK#U.xlsx", usecols=["store_name", "BK#"])

# Merge two tables
merge_uber_last_week = pd.merge (uber_last_week, BKNU_last)
merge_uber_last_week = merge_uber_last_week.astype ({"Sales": int})
# save flirterad file
merge_uber_last_week.to_csv (r'C:files/UberLastWeek!.csv')
clean_uber_last_week = pd.read_csv ("files/UberLastWeek!.csv", usecols=["BK#", "Orders", "Sales", "Unfulfilleds"])
clean_uber_last_week = clean_uber_last_week.rename (
    columns={'Orders': 'Amount of orders (U) Last Week', 'Sales': 'Sales (U) Last Week',
             'Unfulfilleds': 'Denied orders (U) Last Week'})
# save flirterad file
clean_uber_last_week.to_csv (r'C:files/UberLastWeek!.csv')

mergeU_last = pd.merge (clean_uber_last_week, mergeU, on='BK#', how='outer')

# save flirterad file
mergeU_last.to_excel (r'C:files/NEWTABLE!.xlsx')

# ____________________ WOLT this week

wolt = pd.read_excel ('files/WoltWeek.xlsx')
# Map for BK number for Wolt
BKNW = pd.read_excel ("maps\BK#W.xlsx", usecols=["BK#", "Restaurang", "Venue ID"])
# merge map with file
merge_wolt = pd.merge (wolt, BKNW, on="Venue ID")
# save flirterad file to csv
merge_wolt.to_excel (r'files/WoltWeek!.xlsx')
merged_wolt = pd.read_excel ('files/WoltWeek!.xlsx', usecols=['BK#', "Försäljning (SEK)", "Antal beställningar",
                                                              "Antal nekade beställningar"], encoding="utf-8")

merged_wolt['Försäljning (SEK)'] = merge_wolt['Försäljning (SEK)'].str.replace (',', '').astype (float)
merged_wolt = merged_wolt.rename (
    columns={'Antal beställningar': 'Amount of orders (W)', 'Försäljning (SEK)': 'Sales (W)',
             'Antal nekade beställningar': 'Denied orders (W)'})
merged_wolt.to_csv (r'files/WoltWeek!.csv')

mergeW = pd.merge (mergeU_last, merged_wolt, on='BK#', how='outer')
# save flirterad file
mergeW.to_excel (r'C:files/NEWTABLE!.xlsx')

# ____________________ WOLT last week

wolt_last_week = pd.read_excel ('files/WoltLastWeek.xlsx')
# Map for BK number for Wolt
BKNW_last = pd.read_excel ("maps\BK#W.xlsx", usecols=["BK#", "Restaurang", "Venue ID"])
# merge map with file
merge_wolt_last_week = pd.merge (wolt_last_week, BKNW_last, on="Venue ID")
# save flirterad file to csv
merge_wolt_last_week.to_excel (r'files/WoltLastWeek!.xlsx')
merged_wolt_last_week = pd.read_excel ('files/WoltLastWeek!.xlsx',
                                       usecols=['BK#', "Försäljning (SEK)", "Antal beställningar",
                                                "Antal nekade beställningar"], encoding="utf-8")

merged_wolt_last_week['Försäljning (SEK)'] = merge_wolt_last_week['Försäljning (SEK)'].str.replace (',', '').astype (
    float)
merged_wolt_last_week = merged_wolt_last_week.rename (
    columns={'Antal beställningar': 'Amount of orders (W) Last Week', 'Försäljning (SEK)': 'Sales (W) Last Week',
             'Antal nekade beställningar': 'Denied orders (W) Last Week'})
merged_wolt_last_week.to_csv (r'files/WoltLastWeek!.csv')

mergeW_last = pd.merge (mergeW, merged_wolt_last_week, on='BK#', how='outer')
# save flirterad file
mergeW_last.to_excel (r'C:files/NEWTABLE!.xlsx')

# _________________________ FOODORA this week

foodora = pd.read_csv ('files/FoodoraWeek.csv', usecols=["Restaurant ID", "Food value", "Order status"],
                       encoding="utf-8")
# new column for denied orders
foodora['Denied orders'] = foodora['Order status']
# change order status for order value
foodora['Order status'] = foodora['Order status'].map (
    {'Delivered': 1, 'Accepted': 0, 'Cancelled': 0, 'Picked up': 1, ' ': 0})
foodora['Food value'] = np.where ((0 == foodora["Order status"]), 0, foodora["Food value"])

# change denied orders for denied value
foodora['Denied orders'] = foodora['Denied orders'].map (
    {'Delivered': 0, 'Accepted': 1, 'Cancelled': 1, 'Picked up': 0, null: 1})
# Map for BK number
BKNF = pd.read_excel ("maps\BK#F.xlsx", usecols=["BK#", "Restaurant ID"])
BKNF = BKNF.rename (
    columns={'Foodora#': 'Restaurant ID'})
# merge with BK#
merge_foodora = pd.merge (foodora, BKNF, on='Restaurant ID')

# clean for orders
merge_foodora_orders = merge_foodora
delete_row = merge_foodora_orders[merge_foodora_orders["Order status"] == 0].index
merge_foodora_orders = merge_foodora_orders.drop (delete_row)
merge_foodora_orders = merge_foodora_orders[np.isfinite (merge_foodora_orders['Order status'])]

# group columns order status
group_by = merge_foodora_orders.groupby ("BK#")["Order status"].sum ()
group_by1 = merge_foodora_orders.groupby ("BK#")["Food value"].sum ()
mergeF_group = pd.merge (group_by, group_by1, on='BK#')

# group denied column
group_by2 = merge_foodora.groupby ("BK#")["Denied orders"].sum ()
merge_f = pd.merge (mergeF_group, group_by2, on="BK#")

merge_f = merge_f.rename (
    columns={'Order status': 'Amount of orders (F)', 'Food value': 'Sales (F)',
             'Denied orders': 'Denied orders (F)'})

# merge with NEWTABLE
mergeF = pd.merge (mergeW_last, merge_f, on='BK#', how='outer')

# save flirterad file
mergeF.to_excel (r'C:files/NEWTABLE!.xlsx')

# _________________________ FOODORA last week

foodora_last_week = pd.read_csv ('files/FoodoraLastWeek.csv', usecols=["Restaurant ID", "Food value", "Order status"],
                                 encoding="utf-8")
# new column for denied orders
foodora_last_week['Denied orders Last Week'] = foodora_last_week['Order status']
# change order status for order value
foodora_last_week['Order status'] = foodora_last_week['Order status'].map (
    {'Delivered': 1, 'Accepted': 0, 'Cancelled': 0, 'Picked up': 1, ' ': 0})
foodora_last_week['Food value'] = np.where ((0 == foodora_last_week["Order status"]), 0,
                                            foodora_last_week["Food value"])

# change denied orders for denied value
foodora_last_week['Denied orders Last Week'] = foodora_last_week['Denied orders Last Week'].map (
    {'Delivered': 0, 'Accepted': 1, 'Cancelled': 1, 'Picked up': 0, null: 1})
# Map for BK number
BKNF_last_week = pd.read_excel ("maps\BK#F.xlsx", usecols=["BK#", "Restaurant ID"])
BKNF_last_week = BKNF_last_week.rename (
    columns={'Foodora#': 'Restaurant ID'})
# merge with BK#
merge_foodora_last_week = pd.merge (foodora_last_week, BKNF_last_week, on='Restaurant ID')

# clean for orders
merge_foodora_orders_last = merge_foodora_last_week
delete_row_last = merge_foodora_orders_last[merge_foodora_orders_last["Order status"] == 0].index
merge_foodora_orders_last = merge_foodora_orders_last.drop (delete_row_last)
merge_foodora_orders_last = merge_foodora_orders_last[np.isfinite (merge_foodora_orders_last['Order status'])]

# group columns order status
group_by_last = merge_foodora_orders_last.groupby ("BK#")["Order status"].sum ()
group_by1_last = merge_foodora_orders_last.groupby ("BK#")["Food value"].sum ()
mergeF_group_last_week = pd.merge (group_by_last, group_by1_last, on='BK#')

# group denied column
group_by2_last = merge_foodora_last_week.groupby ("BK#")["Denied orders Last Week"].sum ()
merge_f_last = pd.merge (mergeF_group_last_week, group_by2_last, on="BK#")

merge_f_last = merge_f_last.rename (
    columns={'Order status': 'Amount of orders (F) Last Week', 'Food value': 'Sales (F) Last Week',
             'Denied orders Last Week': 'Denied orders (F) Last Week'})

# merge with NEWTABLE
mergeF_last_week = pd.merge (mergeF, merge_f_last, on='BK#', how='outer')

# File with restaurants name
restaurants_name = pd.read_excel ('maps/Master Sm.xlsx', usecols=["BK#", "Restaurant Name", "Franchise company"],
                                  encoding="utf-8")

# merge with Restaurant Name
everything = pd.merge (mergeF_last_week, restaurants_name, on='BK#', how='inner')

# change columns order
column_order = list (everything.columns.values)
everything = everything[['BK#', "Restaurant Name", "Franchise company",
                         'Amount of orders (F)', 'Amount of orders (F) Last Week', 'Sales (F)', 'Sales (F) Last Week',
                         'Denied orders (F)', 'Denied orders (F) Last Week',
                         'Amount of orders (U)', 'Amount of orders (U) Last Week', 'Sales (U)', 'Sales (U) Last Week',
                         'Denied orders (U)', 'Denied orders (U) Last Week',
                         'Amount of orders (W)', 'Amount of orders (W) Last Week', 'Sales (W)', 'Sales (W) Last Week',
                         'Denied orders (W)', 'Denied orders (W) Last Week']]
# sum columns
everything.loc['Totalt'] = everything.select_dtypes (np.number).sum ()

# sum orders
everything["Amount of orders"] = everything[["Amount of orders (F)", "Amount of orders (U)",
                                             "Amount of orders (W)"]].sum (axis=1)
# sum orders last week
everything["Amount of orders Last Week"] = everything[["Amount of orders (F) Last Week",
                                                       "Amount of orders (U) Last Week",
                                                       "Amount of orders (W) Last Week"]].sum (axis=1)
# diffrence between amount in this week and last week
everything['Result Amount'] = everything["Amount of orders"] - everything["Amount of orders Last Week"]
# sum sales
everything["Sales"] = everything[["Sales (F)", "Sales (U)", "Sales (W)"]].sum (axis=1)
# sum sales for last week
everything["Sales Last Week"] = everything[["Sales (F) Last Week", "Sales (U) Last Week",
                                            "Sales (W) Last Week"]].sum (axis=1)
# diffrence between sales in this week and last week
everything['Result Sales'] = everything["Sales"] - everything["Sales Last Week"]
# average for amount
everything['Average amount'] = everything['Amount of orders'] / 7
# !round Average amount
everything['Average amount'] = everything['Average amount'].apply (lambda x: round (x, 2))

# Avg ticket
everything['Avg ticket'] = everything['Sales'] / everything['Amount of orders']
# !round Average amount
everything['Avg ticket'] = everything['Avg ticket'].apply (lambda x: round (x, 2))

# rejected orders
everything["Rejected orders"] = everything[["Denied orders (F)", "Denied orders (U)",
                                            "Denied orders (W)"]].sum (axis=1)
# rejected orders compared to rest
everything["Rejected orders %"] = everything['Rejected orders'] / everything['Amount of orders'] * 100
# !round rejected orders
everything['Rejected orders %'] = everything['Rejected orders %'].apply (lambda x: round (x, 2))

# style __________________________
cell_format.set_bg_color('green')

worksheet.write('A1', 'Ray', cell_format)
# save flirterad file
everything.to_excel (r'C:rapport/Rapport.xlsx')


# delete tables which you don't need

os.remove ('files/WoltLastWeek!.xlsx')
os.remove ('files/WoltWeek!.xlsx')
os.remove ('files/WoltLastWeek!.csv')
os.remove ('files/WoltWeek!.csv')
os.remove ('files/UberLastWeek!.csv')
os.remove ('files/UberWeek!.csv')
os.remove ('files/NEWTABLE!.xlsx')

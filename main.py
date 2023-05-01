# ------------------------------- Specify Parameters ------------------------------- #
path = r"C:\Users\perovj01\Documents\Test\datasets"  # temporary location where files are downloaded (put it between the quotes) -- will be removed at the end of the script
one_drive_path = r"C:\Users\perovj01\OneDrive - etltest\Testing" # where final nonpgi will be saved (put it between the quotes)
# ---------------------------------------------------------------------------------- #
email_subject_1 = "FW: HEI Non PGI" 
email_subject_2 = "FW: HUSA_OrderAllocation_NONPGI_Daily"
email_subject_3 = "FW: Art Scheduler : ART Open Order Report"
email_subject_4 = "FW: Non PGI Report"

import pandas as pd
from datetime import datetime
import os
import shutil
from pathlib import Path
import pyfiglet
from utils import save_attachments, assign_report_date_columns, read_excel_optimized, read_csv_optimized

os.system(f"mkdir {path}")
print(pyfiglet.figlet_format("HUSA NONPGI Version 1.1"))


save_attachments(email_subject_1)
save_attachments(email_subject_2)
save_attachments(email_subject_3)
save_attachments(email_subject_4)

# extract Excel files from the files downloaded
files = os.listdir(path)
filenames = [filename for filename in files if filename.split(".")[1] == "xlsx" or filename.split(".")[1] == "xls"]

# check if we were able to extract Excel files
# script should break if nothing was found
if not filenames:
    raise Exception("No excel files found")

# maersk can be incosistent with headers sometimes with the logo sometimes without 
# this is a vay to circumvet that
try:
    df1 = read_excel_optimized(f"{path}/{filenames[0]}", dataset = 1)
    df1['Balance'] = df1['ORDER QTY'] - df1['ALLOCATED QTY']
except:
    df1 = read_excel_optimized(f"{path}/{filenames[0]}")
    df1['Balance'] = df1['ORDER QTY'] - df1['ALLOCATED QTY']
  
df2 = read_excel_optimized(f"{path}/{filenames[1]}", dataset = 2)
df3 = read_excel_optimized(f"{path}/{filenames[2]}")
df4 = read_excel_optimized(f"{path}/{filenames[3]}")

assign_report_date_columns(email_subject_1, df1)
assign_report_date_columns(email_subject_2, df2)
assign_report_date_columns(email_subject_3, df3)
assign_report_date_columns(email_subject_4, df4)

# ------------------------------- Renames Columns For Consistency ------------------------------- #
df1.rename(columns={
    "WAREHOUSE": "Warehouse",
    "DISTRIBUTOR NAME": "Distributor Name",
    "STATE": "State",
    "DELIVERY ORDER": "Delivery Number",
    "SALES ORDER": "Sales Order Number",
    "DROP DATE": "Order Drop Date",
    "TENDER DATE": "Tender Date",
    "ITEM DESCRIPTION": "Description",
    "ORDER QTY": "Order Qty",
    "ALLOCATED QTY": "Allocated Qty",
    "MAX TENDER DATE": "HUSA Req Tender Date"
}, inplace=True)

df2.rename(columns={
    "Order\nDrop Date": "Order Drop Date",
    "HUSA  Req\nTender Date": "HUSA Req Tender Date",
    "Delivery Order" : "Delivery Number",
    "Sales Order": "Sales Order Number",
    "Hillebrand\nTender Date": "Tender Date"
}, inplace=True)

df3.rename(columns={
    "Wharehouse": "Warehouse",
    "Order No.": "Order Number",
    "Ship To": "Distributor Name",
    "Delivery Order": "Delivery Number",
    "Sales Order": "Sales Order Number",
    "Balance (Short)": "Balance",
    "Item": "HUSA SKU",
    "Request Tender Date": "HUSA Req Tender Date",
    "Order Date": "Order Drop Date",
    "To Ship": "Order Qty"
}, inplace=True)

df4.rename(columns={
    "Wharehouse": "Warehouse",
    "Order-No": "Order Number",
    "Ship To": "Distributor Name",
    "Ship to State": "State",
    "Delivery Order": "Delivery Number",
    "Sales Order": "Sales Order Number",
    "Balance (Short)": "Balance",
    "Item": "HUSA SKU",
    "Request Tender Date": "HUSA Req Tender Date",
    "Order Date": "Order Drop Date",
    "ToShip": "Order Qty"
}, inplace=True)


# ------------------------------- Reorders Columns For Consistency ------------------------------- #
df1 = df1[['Datestamp of Source File','Report Date', 'Warehouse', 'Distributor Name', 'State', 'Delivery Number', 'Sales Order Number'
           , 'Order Drop Date', 'HUSA Req Tender Date', 'Tender Date', 'HUSA SKU', 'Description', 'Order Qty',
           'Allocated Qty', 'Balance']]

df2 = df2[['Datestamp of Source File', 'Report Date', 'Warehouse', 'Distributor Name', 'State', 'Delivery Number', 'Sales Order Number',
           'TMS PO','Order Drop Date', 'HUSA Req Tender Date', 'Tender Date', 'HUSA SKU', 'Description', 'Order Qty',
           'Allocated Qty', 'Balance', 'Type']]

df3['State'] = " "
df3 = df3[['Datestamp of Source File', 'Report Date', 'Warehouse', 'Distributor Name', 'State', 'Delivery Number', 'Sales Order Number',
           'TMS PO','Order Drop Date', 'HUSA Req Tender Date', 'Tender Date', 'HUSA SKU', 'Description', 'Order Qty',
           'Allocated Qty', 'Balance']]

df4 = df4[['Datestamp of Source File', 'Report Date', 'Warehouse', 'Distributor Name', 'State', 'Delivery Number', 'Sales Order Number',
           'TMS PO','Order Drop Date', 'HUSA Req Tender Date', 'Tender Date', 'HUSA SKU', 'Description', 'Order Qty',
           'Allocated Qty', 'Balance']]

combined = pd.concat([df1, df2, df3, df4], ignore_index=True)

now = datetime.now()
current_date_time = now.strftime("%m/%d/%Y %H:%M:%S")
combined['Last Refresh'] = current_date_time

combined = combined[['Last Refresh','Datestamp of Source File', 'Report Date', 'Warehouse', 'Distributor Name', 'State', 'Type', 'Delivery Number', 'Sales Order Number',
           'TMS PO','Order Drop Date', 'HUSA Req Tender Date', 'Tender Date', 'HUSA SKU', 'Description', 'Order Qty',
           'Allocated Qty', 'Balance']]


# move data to onedrive which causes power automate to upload to sharepoint
today = datetime.today().strftime('%m-%d-%Y')
cur_path = Path(f"NON PGI {today}.csv").absolute()

if os.path.isfile(cur_path):
    print("Previous Report Exists -- Appending New Data")
    old_non_pgi = pd.read_csv_optimized(f"NON PGI {today}.csv")
    combined[~combined['Datestamp of Source File'].isin(old_non_pgi['Datestamp of Source File'])]
    newly_created_non_pgi = pd.concat([old_non_pgi, combined], axis=0, ignore_index=True)
    newly_created_non_pgi.to_csv(f"NON PGI {today}.csv", index=False)
else:
    combined.to_csv(f"NON PGI {today}.csv", index=False)

NON_PGI = pd.read_csv(f"NON PGI {today}.csv")
NON_PGI.reset_index(drop=True)

print("Moving to OneDrive")
shutil.copy(cur_path, one_drive_path)
shutil.rmtree(path)

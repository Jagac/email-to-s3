import os
import shutil
from datetime import datetime
from pathlib import Path

import boto3
import pandas as pd
import pyfiglet
import yaml

from utils import Optimization, OutlookDataManipulation

print(pyfiglet.figlet_format("NONPGI V0.2"))

with open("config.yaml", "r") as file:
    global_variables = yaml.safe_load(file)

path = global_variables["path"]
os.mkdir(path)
email = OutlookDataManipulation(path)
oprimizer = Optimization()


def extract():
    global email_subject_1
    global email_subject_2
    global email_subject_3
    global email_subject_4

    email_subject_1 = global_variables["emails"]["email_subject_1"]
    email_subject_2 = global_variables["emails"]["email_subject_2"]
    email_subject_3 = global_variables["emails"]["email_subject_3"]
    email_subject_4 = global_variables["emails"]["email_subject_4"]

    email.save_attachments(email_subject_1)
    email.save_attachments(email_subject_2)
    email.save_attachments(email_subject_3)
    email.save_attachments(email_subject_4)


def load():
    filenames = email.folder_empty_check()

    try:
        df1 = oprimizer.read_excel_optimized(f"{path}/{filenames[0]}", skip_rows=1)
        df1["Balance"] = df1["ORDER QTY"] - df1["ALLOCATED QTY"]
    except:
        df1 = oprimizer.read_excel_optimized(f"{path}/{filenames[0]}")
        df1["Balance"] = df1["ORDER QTY"] - df1["ALLOCATED QTY"]

    df2 = oprimizer.read_excel_optimized(f"{path}/{filenames[1]}", skip_rows=4)
    df3 = oprimizer.read_excel_optimized(f"{path}/{filenames[2]}")
    df4 = oprimizer.read_excel_optimized(f"{path}/{filenames[3]}")

    return df1, df2, df3, df4


def transform(df1, df2, df3, df4):
    df1 = df1.rename(
        columns={
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
            "MAX TENDER DATE": "HUSA Req Tender Date",
        }
    )
    df2 = df2.rename(
        columns={
            "Order\nDrop Date": "Order Drop Date",
            "HUSA  Req\nTender Date": "HUSA Req Tender Date",
            "Delivery Order": "Delivery Number",
            "Sales Order": "Sales Order Number",
            "Hillebrand\nTender Date": "Tender Date",
        }
    )
    df3 = df3.rename(
        columns={
            "Wharehouse": "Warehouse",
            "Order No.": "Order Number",
            "Ship To": "Distributor Name",
            "Delivery Order": "Delivery Number",
            "Sales Order": "Sales Order Number",
            "Balance (Short)": "Balance",
            "Item": "HUSA SKU",
            "Request Tender Date": "HUSA Req Tender Date",
            "Order Date": "Order Drop Date",
            "To Ship": "Order Qty",
        }
    )
    df4 = df4.rename(
        columns={
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
            "ToShip": "Order Qty",
        }
    )

    email.assign_report_date_columns(df1, email_subject_1)
    email.assign_report_date_columns(df2, email_subject_2)
    email.assign_report_date_columns(df3, email_subject_3)
    email.assign_report_date_columns(df4, email_subject_4)

    combined = pd.concat([df1, df2, df3, df4], ignore_index=True)
    now = datetime.now()
    current_date_time = now.strftime("%m/%d/%Y %H:%M:%S")
    combined["Last Refresh"] = current_date_time

    combined = combined[
        [
            "Last Refresh",
            "Datestamp of Source File",
            "Report Date",
            "Warehouse",
            "Distributor Name",
            "State",
            "Type",
            "Delivery Number",
            "Sales Order Number",
            "TMS PO",
            "Order Drop Date",
            "HUSA Req Tender Date",
            "Tender Date",
            "HUSA SKU",
            "Description",
            "Order Qty",
            "Allocated Qty",
            "Balance",
        ]
    ]

    return combined


def main():
    extract()
    df1, df2, df3, df4 = load()
    combined = transform(df1, df2, df3, df4)

    today = datetime.today().strftime("%m-%d-%Y")
    cur_path = Path(f"NON PGI {today}.csv").absolute()

    if os.path.isfile(cur_path):
        old_non_pgi = oprimizer.read_csv_optimized(f"NON PGI {today}.csv")

        combined[
            ~combined["Datestamp of Source File"].isin(
                old_non_pgi["Datestamp of Source File"]
            )
        ]

        newly_created_non_pgi = pd.concat(
            [old_non_pgi, combined], axis=0, ignore_index=True
        )

        newly_created_non_pgi.to_csv(f"NON PGI {today}.csv", index=False)
    else:
        combined.to_csv(f"NON PGI {today}.csv", index=False)

    shutil.rmtree(path)


    s3 = boto3.client('s3')
    bucket = 'non-pgi-emails'
    s3_route = "python/saved_emails"
    filename = f"NON PGI {today}.csv"
    with open(filename, "rb") as file:
        s3.upload_fileobj(filename, bucket, s3_route)

if __name__ == "__main__":
    main()

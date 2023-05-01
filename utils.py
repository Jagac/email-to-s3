from typing import List
import pandas as pd
import win32com.client
from datetime import datetime
import os



def save_attachments(subject_prefix: str):
    """
    Saves whatever attachments (pngs and pdfs included) for the specified subject prefix ie how the email subject starts
    Args:
        subject_prefix: phrase subject starts with
                        omit the date on the email subjects

    Returns: folder with downloaded attachments
    """
    # navigate outlook folders
    path = r"C:\Users\perovj01\Documents\Test\datasets"
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    root_folder = outlook.Folders.Item(1)
    Reports_Folder = root_folder.Folders['Reports']
    NON_PGI_Files_subfolder = Reports_Folder.Folders['NON-PGI Files']
    messages = NON_PGI_Files_subfolder.Items
    
    messages.Sort("[ReceivedTime]", True)
    for message in messages:
        if message.Subject.startswith(subject_prefix):
            print(f"Downloading {message.Subject}")
            for attachment in message.Attachments:
                attachment.SaveAsFile(os.path.join(path, str(attachment.FileName)))
            return

def assign_report_date_columns(subject_prefix : str, df : pd.DataFrame):
    """Assigns report date column based on the current date and assigns a Datestamp column based on timestamp of email received
    Args:
        subject_prefix: phrase subject starts with
                                omit the date on the email subjects
        df:  dataframe that reads downloaded file
             filename and email subject need to match
    """
    # navigate outlook folders
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    root_folder = outlook.Folders.Item(1)
    Reports_Folder = root_folder.Folders['Reports']
    NON_PGI_Files_subfolder = Reports_Folder.Folders['NON-PGI Files']
    messages = NON_PGI_Files_subfolder.Items

    messages.Sort("[ReceivedTime]", True)
    for message in messages:
        if message.Subject.startswith(subject_prefix):
            df['Report Date']= datetime.today().strftime('%m-%d-%Y')
            df['Datestamp of Source File'] = message.Senton
            df['Datestamp of Source File']=df['Datestamp of Source File'].dt.tz_convert(None) # https://stackoverflow.com/questions/51827582/message-exception-ignored-when-dealing-pandas-datetime-type
            return


# ---------------- Newest version saves a bit of memory ---------------- #
def optimize_floats(df: pd.DataFrame):
    floats = df.select_dtypes(include=["float64"]).columns.tolist()
    df[floats] = df[floats].apply(pd.to_numeric, downcast="float")
    
    return df

def optimize_ints(df: pd.DataFrame):
    ints = df.select_dtypes(include=["int64"]).columns.tolist()
    df[ints] = df[ints].apply(pd.to_numeric, downcast="integer")
    
    return df

def optimize_objects(df: pd.DataFrame, datetime_features: List[str]):
    for col in df.select_dtypes(include=["object"]):
        if col not in datetime_features:
            num_unique_values = len(df[col].unique())
            num_total_values = len(df[col])
            if float(num_unique_values) / num_total_values < 0.5:
                df[col] = df[col].astype("category")
        else:
            df[col] = pd.to_datetime(df[col])
            
    return df

def optimize_df(df: pd.DataFrame):
    df = optimize_floats(df)
    df = optimize_ints(df)
    df = optimize_objects(df, [])

    return df


def read_excel_optimized(path: str, dataset = 0):
    if dataset == 1:
        df = pd.read_excel(path, skiprows= 1)
    if dataset == 2:
        df = pd.read_excel(path, skiprows= 4)
    else:
        df = pd.read_excel(path)
    
    return optimize_df(df)

def read_csv_optimized(path: str):
    df = pd.read_csv(path)
    
    return optimize_df(df)


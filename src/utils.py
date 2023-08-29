import os
from datetime import datetime
from typing import List

import pandas as pd
import win32com.client


class OutlookDataManipulation:
    def __init__(self, path) -> None:
        self.path = path

    def outlook_connection(self) -> win32com:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        root_folder = outlook.Folders.Item(1)
        Reports_Folder = root_folder.Folders["Reports"]
        NON_PGI_Files_subfolder = Reports_Folder.Folders["NON-PGI Files"]
        messages = NON_PGI_Files_subfolder.Items
        messages.Sort("[ReceivedTime]", True)

        return messages

    def save_attachments(self, subject_prefix: str) -> None:
        """
        Saves whatever attachments (pngs and pdfs included) for the specified subject prefix ie how the email subject starts
        Args:
            subject_prefix: phrase subject starts with
                            omit the date on the email subjects

        Returns: folder with downloaded attachments
        """
        # navigate outlook folders
        messages = self.outlook_connection()
        for message in messages:
            if message.Subject.startswith(subject_prefix):
                for attachment in message.Attachments:
                    attachment.SaveAsFile(
                        os.path.join(self.path, str(attachment.FileName))
                    )

                return

    def assign_report_date_columns(self, df: pd.DataFrame, subject_prefix: str) -> None:
        """Assigns report date column based on the current date and assigns a Datestamp column based on timestamp of email received
        Args:
            subject_prefix: phrase subject starts with
                                    omit the date on the email subjects
            df:  dataframe that reads downloaded file
                filename and email subject need to match
        """
        # navigate outlook folders
        messages = self.outlook_connection()
        for message in messages:
            if message.Subject.startswith(subject_prefix):
                df["Report Date"] = datetime.today().strftime("%m-%d-%Y")
                df["Datestamp of Source File"] = message.Senton
                df["Datestamp of Source File"] = df[
                    "Datestamp of Source File"
                ].dt.tz_convert(None)

                return

    def folder_empty_check(self):
        files = os.listdir(self.path)

        filenames = [
            filename
            for filename in files
            if filename.split(".")[1] == "xlsx" or filename.split(".")[1] == "xls"
        ]

        # check if we were able to extract Excel files
        # script should break if nothing was found
        if not filenames:
            raise Exception("No excel files found")

        return filenames


# ---------------- Newest version saves a bit of memory ---------------- #
class Optimization:
    def optimize_floats(self, df: pd.DataFrame) -> pd.DataFrame:
        floats = df.select_dtypes(include=["float64"]).columns.tolist()
        df[floats] = df[floats].apply(pd.to_numeric, downcast="float")

        return df

    def optimize_ints(self, df: pd.DataFrame) -> pd.DataFrame:
        ints = df.select_dtypes(include=["int64"]).columns.tolist()
        df[ints] = df[ints].apply(pd.to_numeric, downcast="integer")

        return df

    def optimize_objects(
        self, df: pd.DataFrame, datetime_features: List[str]
    ) -> pd.DataFrame:
        for col in df.select_dtypes(include=["object"]):
            if col not in datetime_features:
                num_unique_values = len(df[col].unique())
                num_total_values = len(df[col])
                if float(num_unique_values) / num_total_values < 0.5:
                    df[col] = df[col].astype("category")
            else:
                df[col] = pd.to_datetime(df[col])

        return df

    def optimize_df(self, df: pd.DataFrame) -> pd.DataFrame:
        df = self.optimize_floats(df)
        df = self.optimize_ints(df)
        df = self.optimize_objects(df, [])

        return df

    def read_excel_optimized(self, path: str, skip_rows=0) -> pd.DataFrame:
        if skip_rows == 1:
            df = pd.read_excel(path, skiprows=1)
        if skip_rows == 4:
            df = pd.read_excel(path, skiprows=4)
        else:
            df = pd.read_excel(path)

        return self.optimize_df(df)

    def read_csv_optimized(self, path: str) -> pd.DataFrame:
        df = pd.read_csv(path)

        return self.optimize_df(df)

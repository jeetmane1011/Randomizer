import pandas as pd
import numpy as np
from openpyxl import load_workbook
from DialogHandler import *

class Workbook:

    def __init__(self, filepath):
        self.filepath = filepath
        try:
            xl = pd.ExcelFile(self.filepath, engine="openpyxl")
            self.sheets = xl.sheet_names
        except Exception as e:
            print(e)

        self.sheets_info = []

    def get_all_names(self):

        df = self.read_sheet_to_df()
        # removing null values and resetting index
        data = df.dropna(axis=0, how='all')
        data = data.fillna("-")
        data = data.reset_index(drop=True)
        self.sheets_info.append([self.sheets[0], data.shape[0], "-"])
        return data

    def get_avail_names(self, data):

        data.reset_index(drop=True, inplace=True)
        all_df = [data]

        columns_to_use = data.columns.values
        for sheet in self.sheets[1:]:
            df = self.read_sheet_to_df(sheet=sheet, usecols=columns_to_use)
            df = df.dropna(axis=0, how="all")
            df = df.fillna("-")
            df = df.reset_index(drop=True)
            candidate_name = self.get_candidate_name(df)
            self.sheets_info.append([sheet, df.shape[0], candidate_name])
            if data.shape[1] == df.shape[1]:
                all_df.append(df)
        
        df_result = pd.concat(all_df, axis=0, ignore_index=True)
        df_result = df_result.drop_duplicates(keep=False, ignore_index=True)
        return df_result

    def get_candidate_name(self, df):
        if df.shape[0] == 0:
            return "-"
        candidate_info = df.iloc[0]
        return '  '.join([str(value) for value in candidate_info])

    def write_to_new_sheet(self, sheet_name, df):
        # write df to excel in a new sheet
        with pd.ExcelWriter(self.filepath, engine='openpyxl', mode='a') as writer:  
            df.to_excel(writer, sheet_name=sheet_name, header=False, index=False)
    
    def read_sheet_to_df(self, sheet=0, usecols=None):
        # read all cells as a string for common format
        return pd.read_excel(self.filepath, sheet_name=sheet, dtype=object, usecols=usecols, header=None, engine="openpyxl")
    

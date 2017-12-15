@echo off & python -x "%~f0" %* & goto :eof

# ==========================================================
# one way to place python script in a batch file
# place python code below (no need for .py file)
# many thanks for the batch file script from jadient:
# https://gist.github.com/jadient/9849314
# ==========================================================

import sys
import pandas as pd
import glob as glob
from pandas.io.excel import ExcelWriter


class MergeTool:
    """
    Merge various files
    Using pandas dataframe
    Save yourself some time
    """

    def __init__(self, merged_sheet_name='_merged_xl_files.xlsx'):
        self.csv_files = None
        self.xlfiles = None
        self.xls = None
        self.merged_sheet_name = merged_sheet_name

    def csv_to_xl(self, globber='*.csv'):
        """
        Combine CSVs
        Into single Excel book
        Each has its own sheet
        """
        self.csv_files = glob.glob(globber)
        with ExcelWriter('_merged_csv_files.xlsx') as ew:
            for csv in self.csv_files:
                pd.read_csv(csv).to_excel(ew, sheet_name=csv[:20], index=False)

    def merge_sheets(self, xlglob='*.xls*'):
        """
        Many to one book
        New sheet names: old book and sheet
        Skip blank dataframes
        """
        self.xlfiles = glob.glob(xlglob)
        with ExcelWriter(self.merged_sheet_name) as ew:
            for wb in self.xlfiles:
                self.xls = pd.read_excel(wb, sheetname=None)
                for name, sheet in self.xls.items():
                    if sheet.shape[0] > 3:
                        sheet.to_excel(
                                ew, (wb[:15].replace(' ', '') +
                                     name[:15].replace(' ', '')), index=False)

x = MergeTool()
x.csv_to_xl()
x.merge_sheets()

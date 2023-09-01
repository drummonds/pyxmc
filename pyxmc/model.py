import arrow
from decimal import Decimal  # If you want to use money this works better

# for floats 0.01 +0.02 != 0.03
import numpy as np
import numpy_financial as npf
from pathlib import Path
from scipy import stats
import shutil

import xlsxwriter
import importlib.metadata


# This is the controlling model which has all the constants in it
#  The various other modules then use this to do
# - capex
# - drug dosage
# - simulation run
class CoreModel:
    def __init__(self):
        self.num_years = 15
        # Defaults for spreadsheet
        self.wb = None
        self.ws = None  # For base case and mean
        self.ws_sd = None  # For Standard Deviation MC
        # Define formats in case not using ws and so not defined later
        self.bold = None
        self.money_fmt = None
        self.bold_money_fmt = None

        self.write_base = True
        try:
            self.version = importlib.metadata.version("pyxmc")
        except importlib.metadata.PackageNotFoundError:
            self.version = "V0.0.0 Alpha"

        self.xlsx_name = "pyxmc.xlsx"
        self.copy_xlsx_path = None

    def create_xlsx(self, filename=None):
        if filename:
            self.xlsx_name = filename
        self.wb = xlsxwriter.Workbook(f"{self.xlsx_name}")
        self.scenario = ""
        self.bold = self.wb.add_format({"bold": 1})
        self.highlight_fmt = self.wb.add_format({"bold": 1, "bg_color": "yellow"})
        self.money_fmt = self.wb.add_format({"num_format": "#,##0"})
        self.bold_money_fmt = self.wb.add_format({"bold": 1, "num_format": "#,##0"})
        # Used on first run to mark where
        self.row = 0

    def close_xlsx(self):
        self.wb.close()
        my_file = Path(self.xlsx_name)
        if self.copy_xlsx_path:
            to_file = Path(self.copy_xlsx_path)
            try:
                shutil.copy(my_file, to_file)
            except PermissionError:
                print("Excel probably open so can't copy over")

    def add_worksheet(self, tabname, colour="green"):
        ws = self.wb.add_worksheet(tabname)
        ws.set_tab_color(colour)
        ws.set_column(0, 0, 25)
        ws.set_column(
            1, self.num_years + 1, 10
        )  # Col B used for single valued data eg NPV
        ws.write_string(0, 0, self.version)
        self.add_years(ws, 1)
        if self.row is None or self.row < 3:
            self.row = 4  # Singleton to mark number of rows used by header
        return ws

    def add_years(self, ws, row):
        ws.write_string(row + 1, 0, "Years", self.bold)
        base_year = arrow.utcnow().shift(years=+1).year
        for i in range(self.num_years):
            ws.write_number(row, i + 2, i + base_year)
            ws.write_number(row + 1, i + 2, i, self.bold)

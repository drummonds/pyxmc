from enum import Enum
from math import sqrt
import numpy as np
from xlsxwriter.utility import xl_range

# class syntax


class VarType(Enum):
    INPUT = 1
    OUTPUT = 2


class AbstractVariable:
    """A variable is a parameter in a model:
    It might be
    - input or output
    - per run (also indexed by 0) or a timeseries index by year
    - single value or stocastic

    The aim of this class is to capture all this information and to use it in a dictionary.
    It is also the aim to capture writing the variable to a spreadsheet report and so at the time
    of capture will write the row the variable has been written at."""

    def __init__(self, core_model, name, title):
        """Clearlines is number of lines to be clear after report"""
        self.cm = core_model
        self.name = name
        self.title = title
        self.clear_lines = 0
        self.row = None
        self.chart_row = None
        self.style = None

    def mark_row(self):
        self.row = self.cm.row  # Default position in spreadsheet for row
        self.cm.row += 1 + self.clear_lines  # Adjust space for amount used

    def scenario_reset(self):
        """Can reset after have done init and mark_row once"""
        pass

    def run_reset(self):
        pass

    def __getitem__(self, key):
        return ""

    def __setitem__(self, key, value):
        raise Exception("Cannot set item value for abstract variable")

    def write_header(self, override_row=None):
        if self.row is None:
            self.mark_row()
        if self.cm.ws_sd is None:  # Base case only
            self._write_header(self.cm.ws, override_row=override_row)
        else:
            self._write_header(self.cm.ws, override_row=override_row)
            self._write_header(self.cm.ws_sd, override_row=override_row)

    def _write_header(self, ws, override_row=None):
        if override_row:
            row = override_row
        else:
            row = self.row
        if self.style:
            ws.write_string(row, 0, self.title, self.style)
        else:
            ws.write_string(row, 0, self.title)

    def write_var(self, override_row=None):
        pass


class Variable(AbstractVariable):
    """A variable is a parameter in a model:
    It might be
    - input or output
    - per run (also indexed by 0) or a timeseries index by year
    - single value or stocastic

    The aim of this class is to capture all this information and to use it in a dictionary.
    It is also the aim to capture writing the variable to a spreadsheet report and so at the time
    of capture will write the row the variable has been written at."""

    def __init__(
        self,
        core_model,
        name,
        title,
        var_type,
        is_single=False,
        is_stochastic=True,
        clear_lines: int = 0,
    ):
        super().__init__(core_model, name, title)
        self.var_type = var_type
        self.row = None  # Row in spreadsheet for result
        self.value = None
        self.is_single = is_single
        self.is_stochastic = is_stochastic
        self.style = None
        self.clear_lines = clear_lines

    def scenario_reset(self):
        """Can reset after have done init and mark_row once"""
        if self.var_type == VarType.INPUT:
            return  # doesn't change
        if self.is_single:
            n = 1
        else:
            n = self.cm.num_years
        self.last_value = np.zeros(n)
        # https://en.wikipedia.org/wiki/Standard_deviation#Rapid_calculation_methods
        # Welford, B. P. (August 1962). "Note on a Method for Calculating Corrected Sums of Squares and Products". Technometrics. 4 (3): 419–420. CiteSeerX 10.1.1.302.7503. doi:10.1080/00401706.1962.10490022
        self.num_samples = np.zeros(n)
        self.A = np.zeros(n)
        self.Q = np.zeros(n)

    def __getitem__(self, key):
        return self.last_value[key]

    def __setitem__(self, key, value):
        value = float(value)
        self.last_value[key] = value
        self.num_samples[key] += 1
        k = self.num_samples[key]
        old_mean = self.A[key]
        if self.title == "Capex" and k > 1.1:
            pass
        self.A[key] += (value - old_mean) / k
        self.Q[key] = self.Q[key] + (value - old_mean) * (value - self.A[key])

    def mean(self, key=0):
        return self.A[key]

    def variance(self, key=0):
        return self.Q[key] / (self.num_samples[key] - 1)

    def stddev(self, key=0):
        return sqrt(self.variance(key=key))

    def num_samples(self, key=0):
        return self.num_samples[key]

    def write_var(self, override_row=None):
        if self.style:
            fmt = self.style
        else:
            fmt = self.cm.money_fmt
        # This is complicated to allow for different cases
        if self.is_single:
            if self.cm.ws_sd is None:  # Base case only
                self.cm.ws.write_number(self.row, 1, self.last_value[0], fmt)
            else:
                self.cm.ws.write_number(self.row, 1, self.mean(0), fmt)
                self.cm.ws_sd.write_number(self.row, 1, self.stdev(0), fmt)
        else:
            for year in range(self.cm.num_years):
                if self.cm.ws_sd is None:  # Base case only
                    self.cm.ws.write_number(
                        self.row, year + 2, self.last_value[year], fmt
                    )
                else:
                    self.cm.ws.write_number(self.row, year + 2, self.mean(year), fmt)
                    self.cm.ws_sd.write_number(
                        self.row, year + 2, self.stddev(year), fmt
                    )

    def add_chart(self, title=None):
        if self.chart_row is None:
            self.chart_row = self.cm.row  # Don't really like this implicit mark
            self.cm.row += 31
        if self.cm.ws_sd is None:  # Base case only
            self._add_chart(self.cm.ws, title=title)
        else:
            self._add_chart(
                self.cm.ws, title=title, error_bars=self.cm.ws_sd.get_name()
            )
            self._add_chart(self.cm.ws_sd, title=title)

    def _add_chart(self, ws, title=None, error_bars=None):
        if title is None:
            title = f"Plot of {self.title}"
        # Create a new chart object.
        chart = self.cm.wb.add_chart({"type": "column"})
        chart.set_title({"name": title})
        chart.set_size({"x_scale": 2, "y_scale": 2})
        chart.set_x_axis({"name": "Years"})
        chart.set_y_axis({"name": "GBP"})
        # Add a series to the chart.
        range = xl_range(self.row, 2, self.row, 1 + self.cm.num_years)
        c = {"name": self.title, "values": f"='{ws.get_name()}'!{range}"}
        if error_bars:
            c["y_error_bars"] = {
                "type": "custom",
                "plus_values": f"='{error_bars}'!{range}",
                "minus_values": f"='{error_bars}'!{range}",
            }
        chart.add_series(c)
        # Insert the chart into the worksheet.
        ws.insert_chart(self.chart_row + 1, 1, chart)


class VariableCopy(AbstractVariable):
    """Copy to appear at different place in spreadsheet"""

    def __init__(self, original: Variable):
        """Clearlines is number of lines to be clear after report"""
        super().__init__(original.cm, original.name, original.title)
        self.original = original
        self.row = None  # Row in spreadsheet for result

    def __getitem__(self, key):
        return self.original[key]

    def __setitem__(self, key, value):
        raise Exception(
            f"Error to set the value of a copy variable name = {self.original.name}"
        )

    def mean(self, key=0):
        return self.original.mean[key]

    def variance(self, key=0):
        return self.original.variance[key]

    def stddev(self, key=0):
        return self.original.stddev[key]

    def num_samples(self, key=0):
        return self.original.num_samples[key]

    def write_header(self, override_row=None):
        self.v.write_header(override_row=self.row)
        if self.cm.ws_sd is None:  # Base case only
            self._write_header(self.cm.ws, override_row=override_row)
        else:
            self._write_header(self.cm.ws, override_row=override_row)
            self._write_header(self.cm.ws_sd, override_row=override_row)

    def write_var(self, override_row=None):
        if override_row:
            row = override_row
        else:
            row = self.row
        if self.style:
            fmt = self.style
        else:
            fmt = self.cm.money_fmt
        # This is complicated to allow for different cases
        if self.is_single:
            if self.cm.ws_sd is None:  # Base case only
                self.cm.ws.write_number(row, 1, self.last_value[0], fmt)
            else:
                self.cm.ws.write_number(row, 1, self.mean(0), fmt)
                self.cm.ws_sd.write_number(row, 1, self.stdev(0), fmt)
        else:
            for year in range(self.cm.num_years):
                if self.cm.ws is None:  # Base case only
                    self.cm.ws.write_number(row, year + 2, self.last_value[year], fmt)
                else:
                    self.cm.ws.write_number(row, year + 2, self.mean(year), fmt)
                    self.cm.ws_sd.write_number(row, year + 2, self.stddev(year), fmt)


class VariableHeader(AbstractVariable):
    """A variable is just an entry a marker in the spreadsheet but acts like other variables"""

    pass


class StringVariable(AbstractVariable):
    """A string variable is time series or single string:
    It might be
    - input or output
    - per run (also indexed by 0) or a timeseries index by year
    """

    def __init__(self, core_model, name, title, var_type, is_single=False):
        super().__init__(core_model, name, title)
        self.var_type = var_type
        self.is_single = is_single

    def scenario_reset(self):
        """Can reset after have done init and mark_row once"""
        if self.var_type == VarType.INPUT:
            return  # doesn't change
        if self.is_single:
            n = 1
        else:
            n = self.cm.num_years
        self.last_value = np.full(n, "", dtype="object")

    def __getitem__(self, key):
        return self.last_value[key]

    def __setitem__(self, key, value):
        self.last_value[key] = value

    def write_var(self, override_row=None):
        if self.style:
            fmt = self.style
        else:
            fmt = self.cm.money_fmt
        # This is complicated to allow for different cases
        if self.is_single:
            if self.cm.ws_sd is None:  # Base case only
                self.cm.ws.write_string(self.row, 1, self.last_value[0], fmt)
            else:
                self.cm.ws.write_string(self.row, 1, self.last_value[0], fmt)
                self.cm.ws_sd.write_string(self.row, 1, self.last_value[0], fmt)
        else:
            for year in range(self.cm.num_years):
                if self.cm.ws_sd is None:  # Base case only
                    self.cm.ws.write_string(
                        self.row, year + 2, self.last_value[year], fmt
                    )
                else:
                    self.cm.ws.write_string(
                        self.row, year + 2, self.last_value[year], fmt
                    )
                    self.cm.ws_sd.write_string(
                        self.row, year + 2, self.last_value[year], fmt
                    )

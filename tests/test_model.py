from pathlib import Path

from pyxmc import CoreModel

from .xlsx_diff import xlsx_same


def test_core():
    cm = CoreModel()
    cm.create_xlsx()
    cm.close_xlsx()
    assert xlsx_same(
        cm.xlsx_name, "./tests/template_xlsx/pyxmc_test_core.xlsx"
    ), "There are differences between generated spreadsheet and template"





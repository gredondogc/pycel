from os.path import abspath, dirname, join
from sys import path

from pycel.excelutil import Cell
from pycel.excelcompiler import ExcelCompiler

if abspath(join(dirname(__file__), '../src')) not in path:
    path.insert(0, join(dirname(__file__), '../src'))

# RUN AT THE ROOT LEVEL
excel = ExcelCompiler(join(dirname(__file__), "../example/example.xlsx")).excel
cursheet = excel.get_active_sheet()


def make_cells():
    global excel, cursheet

    my_input = ['A1', 'A2:B3']
    output_cells = Cell.make_cells(excel, my_input, sheet=cursheet)
    assert len(output_cells) == 3


make_cells()

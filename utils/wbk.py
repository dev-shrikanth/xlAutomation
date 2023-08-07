from openpyxl import Workbook

def init_wb(wb: Workbook):
    wb_calc = wb.calculation
    wb_calc.calcMode = "manual"
    return

def finalize_wb(wb: Workbook):
    wb.calculation.calcMode = "auto"
    return

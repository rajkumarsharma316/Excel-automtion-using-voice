import os
import pythoncom
import win32com.client as win32
import numpy as np

FILE = os.getenv("EXCEL_FILE", "example.xlsx")

excel = None
workbook = None
sheet = None


def connect_excel():
    global excel, workbook, sheet

    if excel and workbook and sheet:
        return

    pythoncom.CoInitialize()

    try:
        excel = win32.GetActiveObject("Excel.Application")
    except:
        excel = win32.Dispatch("Excel.Application")

    excel.Visible = True  # ✅ visible

    path = os.path.abspath(FILE)

    for wb in excel.Workbooks:
        if wb.FullName.lower() == path.lower():
            workbook = wb
            sheet = wb.ActiveSheet
            return

    if os.path.exists(path):
        workbook = excel.Workbooks.Open(path)
    else:
        workbook = excel.Workbooks.Add()
        workbook.SaveAs(path)

    sheet = workbook.ActiveSheet


def write_cell(cell, value):
    connect_excel()
    sheet.Range(cell).Value = value
    workbook.Save()


def delete_cell(cell):
    connect_excel()
    sheet.Range(cell).ClearContents()
    workbook.Save()


def insert_row(row):
    connect_excel()
    sheet.Rows(row).Insert()
    workbook.Save()


def insert_column(col):
    connect_excel()
    sheet.Columns(col).Insert()
    workbook.Save()


def sum_column(col):
    connect_excel()
    last = sheet.Cells(sheet.Rows.Count, col).End(-4162).Row
    result = f"{col}{last+1}"
    sheet.Range(result).Formula = f"=SUM({col}1:{col}{last})"
    workbook.Save()
    return result


def sort_column(col, order="asc"):
    connect_excel()
    last = sheet.Cells(sheet.Rows.Count, col).End(-4162).Row
    sheet.Range(f"{col}1:{col}{last}").Sort(
        Key1=sheet.Range(f"{col}1"),
        Order1=1 if order == "asc" else 2,
        Header=1
    )
    workbook.Save()


def format_bold(col):
    connect_excel()
    sheet.Columns(col).Font.Bold = True
    workbook.Save()


def filter_values(col, condition):
    connect_excel()
    sheet.UsedRange.AutoFilter(
        Field=sheet.Range(f"{col}1").Column,
        Criteria1=condition
    )
    workbook.Save()


def create_chart(x_col, y_col):
    connect_excel()
    xlLine = 4
    chart = workbook.Charts.Add()
    chart.ChartType = xlLine

    xlUp = -4162
    x_last = sheet.Cells(sheet.Rows.Count, x_col).End(xlUp).Row
    y_last = sheet.Cells(sheet.Rows.Count, y_col).End(xlUp).Row

    chart.SetSourceData(
        sheet.Range(f"{x_col}1:{x_col}{x_last},{y_col}1:{y_col}{y_last}")
    )
    chart.Location(2, sheet.Name)
    workbook.Save()


def run_regression(x_col, y_col):
    connect_excel()

    xlUp = -4162
    last = max(
        sheet.Cells(sheet.Rows.Count, x_col).End(xlUp).Row,
        sheet.Cells(sheet.Rows.Count, y_col).End(xlUp).Row
    )

    X, Y = [], []
    for r in range(1, last+1):
        xv = sheet.Range(f"{x_col}{r}").Value
        yv = sheet.Range(f"{y_col}{r}").Value
        if xv and yv:
            try:
                X.append(float(xv))
                Y.append(float(yv))
            except:
                pass

    X = np.array(X)
    Y = np.array(Y)

    m, b = np.polyfit(X, Y, 1)
    r2 = 1 - np.sum((Y-(m*X+b))**2) / np.sum((Y-Y.mean())**2)

    sheet.Range("Z1").Value = "Regression"
    sheet.Range("Z2").Value = "Slope"
    sheet.Range("Z3").Value = "Intercept"
    sheet.Range("Z4").Value = "R²"

    sheet.Range("AA2").Value = m
    sheet.Range("AA3").Value = b
    sheet.Range("AA4").Value = r2

    workbook.Save()

import os
from os import listdir
from os.path import isfile, join
from win32com.client import Dispatch
import linecache
import sys
import com_wrapper
from enum import Enum

# Current working directory
cwd = os.getcwd()

# Global vars
RESULTS_XLSX = cwd + r'\results.xlsx'
REPORTS_DIR = cwd + r'\reports'


class OrderEnum(Enum):
    ALPHABETICAL = 1
    FORWARDS = 2
    BACKWARDS = 3


def PrintException():
    exc_type, exc_obj, tb = sys.exc_info()
    f = tb.tb_frame
    lineno = tb.tb_lineno
    filename = f.f_code.co_filename
    linecache.checkcache(filename)
    line = linecache.getline(filename, lineno, f.f_globals)
    print('EXCEPTION IN ({}, LINE {} "{}"): {}'.format(filename, lineno, line.strip(), exc_obj))
    return


def concat_sheets():

    try:
        _xl = Dispatch('Excel.Application')
        # Calling xl wrapper which retries attempts when calls fail
        xl = com_wrapper.ComWrapper(_xl)
        xl.DisplayAlerts = False
        xl.Visible = True  # You can remove this line if you don't want the Excel application to be visible
        wb_target = xl.Workbooks.Add()
    except Exception:
        print("Error: Failed to open new workbook - close all open workbooks")
        PrintException()
        return False

    try:
        # Get files in REPORTS_DIR
        client_files = [f for f in listdir(REPORTS_DIR) if isfile(join(REPORTS_DIR, f))]
    except Exception:
        print("Error: Failed to get filenames in reports directory")
        PrintException()
        return False

    try:
        for idx, f in enumerate(client_files):
            path_source = os.path.join(REPORTS_DIR, f)
            print("path_source = " + path_source)
            wb_source = xl.Workbooks.Open(Filename=path_source)
            ws_source = wb_source.Worksheets(1)
            client_name = str(ws_source.Range("A2").value)
            print("client = " + client_name)
            ws_source.Name = client_name
            insert_worksheet_ordered(wb_target, ws_source, OrderEnum.ALPHABETICAL)
            wb_source.Close()
    except PermissionError:
        print("Error: Workbook needs to be closed first")
        return False
    except Exception:
        print("Error: Workbook unknown")
        PrintException()
        return False

    if len(client_files) > 0:
        workbook_clean_save(wb_target)

    xl.DisplayAlerts = True
    xl.Quit()

    return True


def insert_worksheet_ordered(wb_target, ws_source, enum):

    num_sheets = wb_target.Worksheets.count

    if enum == OrderEnum.ALPHABETICAL:
        for ws_target in wb_target.Worksheets:
            if ws_source.Name < ws_target.Name:
                ws_source.Copy(Before=ws_target)
                break
        else:  # Only if for loop reached end without break
            ws_source.Copy(Before=None, After=wb_target.WorkSheets(num_sheets))
    elif enum == OrderEnum.FORWARDS:
        ws_source.Copy(Before=None, After=wb_target.WorkSheets(num_sheets))
    elif enum == OrderEnum.BACKWARDS:
        ws_source.Copy(Before=wb_target.Worksheets(1))
    else:
        print("Warning: Sheet added to start of workbook - enum.value = {}, value invalid".format(enum.value))
        ws_source.Copy(Before=wb_target.Worksheets(1))

    return True


def workbook_clean_save(wb):

    # Delete Sheet1 that's auto-created with new workbook
    if wb.Sheets.count > 1:
        try:
            wb.Worksheets('Sheet1').Delete()
        except Exception:
            print("Error: Failed to delete Sheet1")
            PrintException()

    # First delete an existing results file
    if isfile(RESULTS_XLSX):
        try:
            os.remove(RESULTS_XLSX)
        except Exception:
            print("Error: Failed to delete current results.xlsx")
            PrintException()

    try:
        wb.SaveAs(RESULTS_XLSX)
    except Exception:
        print("Error: Failed to save results.xlsx")
        PrintException()
        return False

    return True


def main():

    print("START - FILES HANDLING")

    print("cwd = {}".format(cwd))

    # Amend results.xlsx
    concat_sheets()

    print("END - FILES HANDLING")

    return True


if __name__ == '__main__': main()

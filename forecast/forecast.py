#!/usr/bin/python3
# Version 1.0
# Sean McGlincy
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
from datetime import datetime
from datetime import timedelta
import xlwt 
# from xlutils.copy import copy
    

def map_employees(df, column_name):
    employee_map = {}
    employee_list = df.loc[:, column_name]
    n = len(employee_list)
    for i in range(0, n):
        employee_map[employee_list[i]] = i
    return employee_map

def get_times(df, row_number, project_type):
    return df.loc[row_number, project_type]

def employee_hash(df_employee, col):
    employee_hash = {}
    for index, row in  df_employee[[col.employee]].iterrows():
        employee = row[col.employee]
        employee_hash[employee] = float(0.0)
    return employee_hash

def employee_list(df_employee, col):
    employee_list = []
    for index, row in  df_employee[[col.employee]].iterrows():
        employee = row[col.employee]
        employee_list.append(employee)
    employee_list.sort()
    return employee_list

def str_to_date(str):
    return datetime.strptime(str, '%m/%d/%Y')

def date_to_str(dt):
    return datetime.strftime(dt, '%m/%d/%Y')

def next_monday(dt):
    # 0 = Monday, 1=Tuesday, 2=Wednesday...
    return  dt + timedelta(days=(7 - dt.weekday()))
def get_all_dates(hash):
    date_list = [] 
    for key in hash.keys():
        date_list.append(str_to_date(key))
    date_list.sort()
    return  [date_to_str(date) for date in date_list]

def forecast(df_project, df_employee, col,completed_flag):
    # Hash monday dates.  date = subhash {employee: hour_sum}
    forecast = {}
    # Map employee row numbers for lookup
    planner_map = map_employees(df_employee, col.employee)

    for index, row in df_project.iterrows():
        planner = row[col.planner]
        project_type  = row[col.project_type]
        date = row[col.date]

        # TODO check date, project_type and planer

        # get next monday date
        date = date_to_str(next_monday(date))
        if date not in forecast:
            forecast[date] = employee_hash(df_employee, col)

        # Hours
        row_num = int(planner_map[planner])
        hours = get_times(df_employee, row_num, project_type)

        if completed_flag:
            if row[col.status] == col.complete:
                hours = 0
        forecast[date][planner] = forecast[date][planner] + hours
    return forecast


def read_input(filename, sheet, col):
    # Read Excel and import as dataframe 'df'
    df_project = pd.read_excel(filename, sheet_name=sheet.project)
    df_employee = pd.read_excel(filename, sheet_name=sheet.employee)

    # Convert Column names to lowercase
    df_project.columns = df_project.columns.str.lower()
    df_employee.columns = df_employee.columns.str.lower()

    # Convert columns to lowercase
    df_project[col.planner] = df_project[col.planner].str.lower() 
    df_project[col.project_type] = df_project[col.project_type].str.lower() 
    df_project[col.status] = df_project[col.status].str.lower() 
    df_employee[col.employee] = df_employee[col.employee].str.lower()

    return df_project, df_employee

def write_output(book, sheetname, col, results, list_employees):

    sh = book.add_sheet(sheetname)
    style_dec = xlwt.XFStyle()
    style_dec.num_format_str = '0.00'

    list_dates = get_all_dates(results)

    sh.write(0, 0, 'Employees')
    column = 0
    row = 1
    for employee in list_employees:
        sh.write(row, column, employee.title())
        row += 1

    for j in range(0, len(list_dates)):
        row = 0 
        column = j + 1
        date = list_dates[j]
        sh.write(row, column, date)
        r = results[date]
        
        for employee in list_employees:
            row += 1
            sh.write(row, column, r[employee], style_dec)

# To Do Check that employee column matches Project
def main():
    """ Main program """
    # Excel Files
    filename = 'Work Load Forecast.xlsx'

    # Excel Sheets
    sheet = lambda:0
    sheet.project = 'Project Log'
    sheet.employee = 'Employee Times'

    # Column Headers
    # Lowercase
    col = lambda:0
    col.date = 'date'
    col.project_type = 'project type'
    col.planner = 'planner'
    col.status = 'status'
    col.employee = "employee"
    col.complete = 'complete'
    col.pending = 'pending'

    # Returns data frames for project and employee 
    df_project, df_employee = read_input(filename, sheet, col)
    list_employees = employee_list(df_employee, col)

    # Forcast
    results = forecast(df_project, df_employee, col, False)
    results_exclude_completed = forecast(df_project, df_employee, col, True)

    # FileName
    todays_date = datetime.strftime(datetime.today(), '%Y-%m-%d')
    report_name = 'report-' + todays_date + '.xlsx'

    # Write Data
    book = xlwt.Workbook()
    write_output(book, 'total-' + todays_date, col, results, list_employees)
    write_output(book, 'exclude-completed' + todays_date, col, results_exclude_completed, list_employees)
    book.save(report_name)

    return 0

if __name__ == "__main__":
    main()
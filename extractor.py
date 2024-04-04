import pandas as pd
from datetime import timedelta
from math import ceil
from xlsxwriter.utility import xl_range
import numpy as np
def process_excel_files(data_file_path, data_sheet_name, employee_file_path, employee_sheet_name, output_file_path):
   
    statuses = ['Approved', 'Missing', 'Working', 'Submitted', 'Rejected', 'Error']
    status_colors = {'Approved': '#C6EFCE', 'Missing': '#FFFFFF', 'Working': '#FFC7CE', 'Submitted': '#FFFF00',  'Rejected': '#FF0000',  'Error': '#0000FF'  }
    df = pd.read_excel(data_file_path, sheet_name=data_sheet_name, header=14)
    df = df[df['Project Number'] != '-'].copy()
    employees_df = pd.read_excel(employee_file_path, sheet_name=employee_sheet_name, header=0)
   
    date_format = '%d-%b-%Y' 
    def get_week_of_month(dt):
        first_day = dt.replace(day=1)
        dom = dt.day
        adjusted_dom = dom + first_day.weekday()
        return int(ceil(adjusted_dom/7.0))

    def reformat_name(name):
        titles = {'Miss', 'Mr.', 'Mrs.', 'Dr.', 'Ms.'}
        parts = name.replace(',', '').split()
        filtered_parts = [part for part in parts if part not in titles]
        return ' '.join(filtered_parts).upper()

    df['Employee Full Name'] = df['Employee Name'].apply(reformat_name)
    df['Employees'] = df['Employee Number']+ ' - ' + df['Employee Full Name']

    df['Time Period Start'] = pd.to_datetime(df['Time Period Start'], format=date_format)
    df['Time Period End'] = pd.to_datetime(df['Time Period End'], format=date_format)
    df['Projects'] = df['Project Number'].astype(str) + ' - ' + df['Project Name']
    df['Year'] = df['Time Period Start'].dt.year
    df['Month'] = df['Time Period Start'].dt.month
    df['Week_of_Month'] = df['Time Period Start'].apply(get_week_of_month)
    df['Hours'] = pd.to_numeric(df['Hours'], errors='coerce')
    
    employee_full_names = employees_df['SL TEAM'].apply(reformat_name).tolist()
    filtered_df = df[df['Employee Full Name'].isin(employee_full_names)].copy()
    filtered_df['Week_of_Year'] = filtered_df['Time Period Start'].dt.isocalendar().week

    status_df = filtered_df.groupby(['Projects', 'Employees', 'Year', 'Week_of_Year', 'Status'])['Hours'].sum().unstack(fill_value=0).reset_index()

    pivot_df = filtered_df.pivot_table(index=['Projects', 'Employees'], columns=['Year', 'Week_of_Year'], values='Hours', aggfunc='sum', fill_value=0)
    pivot_df.columns = [' '.join([str(v) for v in col]).strip() if isinstance(col, tuple) else col for col in pivot_df.columns]
    pivot_df.reset_index(inplace=True)
   
    writer = pd.ExcelWriter(output_file_path, engine='xlsxwriter')
    pivot_df.to_excel(writer, sheet_name='Sheet1', index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    
    merge_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
    header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
    legend_start_row = len(pivot_df) + 2
    for i, (status, color) in enumerate(status_colors.items()):
        worksheet.write_string(legend_start_row + i, 0, status, 
                            workbook.add_format({'bg_color': color, 'bold': True}))
    for col_num, value in enumerate(pivot_df.columns.values):
        worksheet.write(0, col_num, value, header_format)

    status_format = {status: workbook.add_format({'bg_color': color}) for status, color in status_colors.items()}
    present_statuses = [status for status in statuses if status in status_df.columns]

    for row_idx, pivot_row in pivot_df.iterrows():
        employee_name = pivot_row['Employees']
        project_group = pivot_row['Projects']
        for col_name in pivot_df.columns[2:]:
            year, week = map(int, col_name.split()[-2:])
            hours_value = pivot_row[col_name]
            status_row = status_df[(status_df['Projects'] == project_group) & (status_df['Employees'] == employee_name) & (status_df['Year'] == year) & (status_df['Week_of_Year'] == week)]
            if not status_row.empty:
                status_applied = False  
                for status in present_statuses:
                    if status in status_row.columns:
                        status_hours = status_row[status].iloc[0]
                        if status_hours > 0:
                            col_idx = pivot_df.columns.get_loc(col_name)
                            worksheet.write(row_idx + 1, col_idx, hours_value, status_format[status])
                            status_applied = True
                            break  
                if not status_applied:
                    col_idx = pivot_df.columns.get_loc(col_name)
                    worksheet.write(row_idx + 1, col_idx, 0 if hours_value > 0 else "-", merge_format)
            else:
                col_idx = pivot_df.columns.get_loc(col_name)
                worksheet.write(row_idx + 1, col_idx, "-", merge_format)
    start_row = 1

    for row_num in range(1, len(pivot_df) + 1):  
        project_group = pivot_df.at[row_num - 1, 'Projects'] if row_num - 1 < len(pivot_df) else ""
        next_project_group = pivot_df.at[row_num, 'Projects'] if row_num < len(pivot_df) else ""
        if project_group != next_project_group and row_num > start_row:
            worksheet.merge_range(start_row, 0, row_num, 0, project_group, merge_format)
            start_row = row_num + 1
    worksheet.set_column('A:A', 50)  # Assuming 'Project Group' is in column B
    worksheet.set_column('B:B', 50) 
    writer.close()


import os
import openpyxl
import statistics
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.axis import DateAxis
from openpyxl.styles import Border, Font, Side


# returns a tuple with start row in 0 index, end row in 1
def find_start_end(sheet, year):
    found = False
    start_row = 0
    end_row = 0
    for col in sheet.iter_cols(max_col=1, values_only=True):
        for val in col:
            try:
                end_row += 1
                if not found:   # only adjust start if row not found
                    start_row += 1
                if val.strftime("%Y") == str(year):
                    found = True
                elif found:
                    end_row -= 1
                    break
            except AttributeError:    # if we hit a string
                pass
    return start_row, end_row


# returns a dictionary containing all "data types" (i.e. PH) and
# their corresponding column number
def get_data_types(sheet, row_num):
    data_dict = {}
    curr_row = 0
    for row in sheet.iter_rows(min_row=row_num, max_row=row_num, values_only=True):
        for val in row:
            if val is not None:
                curr_row += 1
                data_dict[val.upper()] = curr_row
    return data_dict


def is_number(num):
    try:
        float(num)
        return True
    except (ValueError, TypeError):
        return False


def transfer_data(doc_wb, write_wb, sheet_name, num_sheets, wb_dir, year):
    os.chdir(wb_dir)
    write_sheet = write_wb.create_sheet(sheet_name + ' COMPILED')
    data_types = get_data_types(doc_wb[sheet_name + str(1)], 4)
    dt_keys = list(data_types.keys())
    start_end = []
    date_range = 0
    bold = Font(bold=True)
    for sheet_num in range(1, num_sheets+1):
        doc_sheet = doc_wb[sheet_name + str(sheet_num)]  # i.e. KARP SYSTEM1
        start_end = find_start_end(doc_sheet, year)
        date_range = start_end[1] - start_end[0] + 1
        if sheet_num == 1:   # add the date column the first time around
            start = start_end[0]
            write_sheet['A1'].value = 'DATE'
            write_sheet['A1'].font = bold
            write_sheet['B1'].value = 'DAY'
            write_sheet['B1'].font = bold
            for i in range(date_range):  # number of readings in a year
                write_sheet['A' + str(i + 2)].value = \
                    doc_sheet['A' + str(i + start)].value.strftime('%m/%d/%Y').lstrip('0').replace('/0', '/')
                write_sheet['B' + str(i + 2)].value = doc_sheet['B' + str(i + start)].value
        # add all data into the sheet
        for data_key in dt_keys[2:]:  # add all types of data to row 1
            col_num = data_types[data_key] - 3  # -3 because of date and day columns, i.e. 1st data col will = 0
            c = write_sheet.cell(row=1, column=3+col_num*(num_sheets + 2))  # write the column headers
            c.value = data_key
            c.font = Font(bold=True)
            if col_num != 0:  # so we don't have a random border
                c.border = Border(left=Side(border_style='medium'))
            mean_hc = write_sheet.cell(row=1, column=3 + num_sheets + col_num * (num_sheets + 2))
            mean_hc.value = 'MEAN'
            mean_hc.font = bold
            avg_mean_hc = write_sheet.cell(row=1, column=4 + num_sheets + col_num * (num_sheets + 2))
            avg_mean_hc.value = 'AVG. MEAN'
            avg_mean_hc.font = bold
            for i in range(date_range):  # copy over data
                write_sheet.cell(row=i+2, column=(num_sheets + 2)*col_num+sheet_num+2).value = \
                    doc_sheet.cell(row=start_end[0]+i, column=data_types[data_key]).value

    # now to find the mean of all readings
    for row_num in range(2, date_range+2):  # get the proper row numbers
        for data_key in dt_keys[2:]:
            col_num = data_types[data_key] - 3
            curr_col = col_num * (num_sheets + 2) + 3
            mean = 0
            count = 0  # needed as some columns contain None or a string
            for row in write_sheet.iter_rows(min_row=row_num, max_row=row_num,
                                             min_col=curr_col, max_col=curr_col+num_sheets-1,
                                             values_only=True):
                for val in row:
                    if is_number(val):
                        count += 1
                        mean += float(val)
            if count != 0:  # if count == 0, then we know the mean of these values is not legitimate
                write_sheet.cell(row=row_num, column=curr_col+num_sheets).value = round(mean/count, 6)

    # finding the average mean
    for data_key in dt_keys[2:]:
        col_num = data_types[data_key] - 3
        avg_col = col_num * (num_sheets + 2) + num_sheets + 4  # this is the column which will have average of means
        mean_col = avg_col - 1  # this is the column with means in it
        avg_mean = 0
        all_means = []
        count = 0
        for repeat in range(2):  # very similar code ran twice so we can just change it slightly using conditionals
            for row_num in range(2, date_range+2):
                if not repeat:  # still looking for average means
                    curr_mean = write_sheet.cell(row=row_num, column=mean_col).value
                    if is_number(curr_mean):
                        count += 1
                        avg_mean += curr_mean
                        all_means.append(curr_mean)
                else:  # now actually populate the column with average means
                    curr_cell = write_sheet.cell(row=row_num, column=avg_col)
                    if count != 0:
                        curr_cell.value = round(avg_mean/count, 6)
                    else:
                        curr_cell.value = None
                    curr_cell.border = Border(right=Side(border_style='medium'))
        if all_means:  # if our array isn't empty
            write_sheet.cell(row=date_range+2, column=mean_col).value = 'STDEV:'
            write_sheet.cell(row=date_range+2, column=mean_col).font = bold
            write_sheet.cell(row=date_range+2, column=avg_col).value = statistics.pstdev(all_means)

        c = LineChart()
        dates = Reference(worksheet=write_sheet, max_col=1, min_row=2, max_row=start_end[1]-start_end[0]+2)
        c.title = data_key.upper()
        c.style = 12
        c.x_axis = DateAxis()
        c.x_axis.number_format = 'm/d/yyyy'
        c.x_axis.majorTimeUnit = 'months'
        c.x_axis.title = 'Date'

    write_wb.save('writebook.xlsx')


def run():
    os.chdir('c:\\Users\\aakru\\Downloads')
    doc1 = openpyxl.load_workbook('file 1 user input values waterqual.xlsx')
    write_book = openpyxl.Workbook()
    transfer_data(doc1, write_book, 'KARP SYS', 6, 'c:\\Users\\aakru\\Downloads', 2018)


run()
